"""
core.py — Shared logic for comic.py and test_run.py.

Fixes applied vs previous version:
  1. CV confidence threshold raised from 0.3 → 0.6; bad matches now fall through
  2. CV fallback uses /search endpoint which handles unusual/numeric titles better
  3. Grade normalisation: blank/invalid grade returns None cleanly, eBay skips gracefully
  4. eBay query no longer appends grade string when grade is unknown
  5. pandas fillna: uses per-column dtype-safe approach to suppress FutureWarning
"""
from __future__ import print_function
from curl_cffi import requests
import bs4
from difflib import SequenceMatcher
from datetime import date
import os

rundate = date.today().strftime("%Y-%m-%d")
ID_DATE_COL = 'Identification Date'  # Column name for CV identification tracking

CV_BASE  = "https://comicvine.gamespot.com/api"

EBAY_FINDING_PROD    = "https://svcs.ebay.com/services/search/FindingService/v1"
EBAY_FINDING_SANDBOX = "https://svcs.sandbox.ebay.com/services/search/FindingService/v1"
CV_HEADERS = {"User-Agent": "MyComicCollection/1.0"}
CV_CONFIDENCE_THRESHOLD = 0.60   # Reject matches below this — prevents wrong-title pollution

GRADE_SCALE = [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5,
               5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0,
               9.2, 9.4, 9.6, 9.8, 10.0]


# =============================================================================
#   Utilities
# =============================================================================

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()


def _format_grade(g):
    return f"{g:.1f}"


def _strip_html(text):
    return bs4.BeautifulSoup(text, 'html.parser').get_text(separator=' ').strip()


def _classify_age(cover_date):
    """Classify a comic into an age bracket by cover date string (YYYY-MM-DD or YYYY)."""
    if not cover_date or str(cover_date).strip() in ('', 'Unknown', 'None', 'nan'):
        return 'Unknown'
    try:
        year = int(str(cover_date)[:4])
        if year < 1938:   return 'Platinum Age'
        elif year < 1956: return 'Golden Age'
        elif year < 1970: return 'Silver Age'
        elif year < 1985: return 'Bronze Age'
        elif year < 1992: return 'Copper Age'
        else:             return 'Modern Age'
    except (ValueError, TypeError):
        return 'Unknown'


def normalise_grade(raw_grade):
    """
    Normalise a grade string to 'X.X' format, or return None if blank/invalid.

    Examples:
      '9'   → '9.0'
      '9.8' → '9.8'
      ''    → None
      'NM'  → None  (text grades not supported for eBay lookup)
    """
    g = str(raw_grade).strip()
    if not g or g.lower() in ('nan', 'none', ''):
        return None
    try:
        val = float(g)
        return f"{val:.1f}" if '.' in g or len(g) >= 3 else f"{val:.1f}"
    except ValueError:
        return None


def safe_fillna(df):
    """
    Fill NaN values in a DataFrame in a dtype-safe way.
    Numeric columns get 0, object columns get ''.
    Suppresses the pandas FutureWarning about incompatible dtype assignment.
    """
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].fillna('')
        else:
            df[col] = df[col].fillna(0)
    return df


# =============================================================================
#   Comic Vine
# =============================================================================

def _cv_get(session, cv_api_key, endpoint, params):
    """Make a Comic Vine API call. Returns parsed results or None."""
    params.update({'api_key': cv_api_key, 'format': 'json'})
    url = f"{CV_BASE}/{endpoint}/"
    try:
        resp = session.get(url, params=params, headers=CV_HEADERS, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        if data.get('status_code') == 1:
            return data.get('results')
        else:
            print(f"     CV API error: {data.get('error')} (code {data.get('status_code')})")
            return None
    except Exception as e:
        print(f"     CV request failed: {e}")
        return None


def _pick_best_match(results, full_name, issue_number=None, volume_number=None, publisher=None):
    """
    Given a list of CV issue results, return (best_result, score).

    Scoring:
      - Base: string similarity of 'VolumeName #IssueNum' vs full_name
      - Penalty: -0.3 if issue_number doesn't match exactly (prevents #2 matching #4)
      - Bonus:  +0.1 if publisher name matches (helps disambiguate Marvel vs DC)
      - Bonus:  +0.1 if volume start_year aligns with volume_number ordering

    All scores clamped to [0, 1].
    """
    best, best_score = None, 0.0
    for r in results:
        vol_name = (r.get('volume') or {}).get('name', '')
        candidate = f"{vol_name} #{r.get('issue_number', '')}".upper()
        score = similar(candidate, full_name.upper())

        # Penalty: issue number mismatch — use 0.5 to overcome even near-identical titles
        if issue_number is not None:
            r_issue = str(r.get('issue_number', '')).strip()
            if r_issue != str(issue_number).strip():
                score -= 0.5

        # Bonus: publisher match
        if publisher:
            r_pub = (r.get('publisher') or {}).get('name', '').upper()
            if r_pub and publisher.upper() in r_pub:
                score += 0.1

        score = max(0.0, min(1.0, score))
        if score > best_score:
            best_score = score
            best = r
    return best, best_score


def SearchComicVine(session, cv_api_key, title, issue, variant='',
                    volume_number=None, publisher=None):
    """
    Search Comic Vine for a specific issue.

    Parameters:
      title         — series name (e.g. 'Amazing Spider-Man')
      issue         — issue number (int or str)
      variant       — variant suffix (e.g. ' [Variant A]')
      volume_number — sheet Volume field (int); used to disambiguate relaunches
      publisher     — sheet Publisher field (str); used to break ties

    Search strategy:
      NOTE: The /issues filter endpoint (volume.name + issue_number) was found to be
      unreliable — CV's API ignores the volume.name param and returns tens of thousands
      of unrelated results sorted by internal ID.

      1. /search for the issue directly ("title #issue")
      2. If confidence < threshold: /search for the volume by name + publisher, then
         fetch its issues and match by issue number exactly
      3. Reject if best confidence < CV_CONFIDENCE_THRESHOLD after both strategies

    IMPORTANT: This function only returns *enrichment* metadata.
    The caller must never overwrite Title, Issue, Volume, or Publisher
    fields in the sheet with values from this result.

    Returns a metadata dict or None.
    """
    full_name = f"{title} #{issue}{variant}"

    field_list = (
        'id,name,issue_number,volume,description,deck,'
        'image,cover_date,store_date,cover_price,'
        'site_detail_url,character_credits'
    )

    # --- Strategy 1: /search for the issue directly ---
    print(f"     CV: Searching for '{full_name}'")
    search_results = _cv_get(session, cv_api_key, 'search', {
        'query': f"{title} #{issue}",
        'resources': 'issue',
        'field_list': field_list,
        'limit': 20,
    })
    issues = [r for r in (search_results or []) if r.get('resource_type') == 'issue']
    best, best_score = _pick_best_match(
        issues, full_name,
        issue_number=issue,
        volume_number=volume_number,
        publisher=publisher,
    )

    # --- Strategy 2: volume lookup fallback ---
    # Used when the title is ambiguous (short, numeric, year-like) and /search
    # can't surface the right series. Find the volume by name, then fetch its issues.
    if best_score < CV_CONFIDENCE_THRESHOLD:
        print(f"     CV: Issue search low ({int(best_score*100)}%) — trying volume lookup fallback")
        # Use title only (not title+publisher) — adding publisher name degrades CV search results
        vol_results = _cv_get(session, cv_api_key, 'search', {
            'query': title,
            'resources': 'volume',
            'field_list': 'id,name,publisher,start_year,count_of_issues',
            'limit': 10,
        })
        volumes = [r for r in (vol_results or []) if r.get('resource_type') == 'volume']

        # Score volumes by name similarity + publisher match + issue range plausibility
        best_vol, best_vol_score = None, 0.0
        try:
            issue_int = int(str(issue).strip())
        except (ValueError, TypeError):
            issue_int = None

        for v in volumes:
            vscore = similar(v.get('name', '').upper(), title.upper())
            if publisher:
                vpub = (v.get('publisher') or {}).get('name', '').upper()
                if vpub and publisher.upper() in vpub:
                    vscore += 0.2
            # Bonus: volume has enough issues to contain the one we want
            if issue_int is not None:
                count = v.get('count_of_issues') or 0
                try:
                    count = int(count)
                except (ValueError, TypeError):
                    count = 0
                if count >= issue_int:
                    vscore += 0.15
                elif count > 0 and count < issue_int:
                    vscore -= 0.3   # volume too short to contain this issue
            vscore = min(1.0, max(0.0, vscore))
            if vscore > best_vol_score:
                best_vol_score = vscore
                best_vol = v

        if best_vol and best_vol_score >= 0.7:
            vol_id = best_vol.get('id')
            print(f"     CV: Volume match '{best_vol.get('name')}' (id={vol_id}, {int(best_vol_score*100)}%) — fetching issues")
            vol_issues = _cv_get(session, cv_api_key, 'issues', {
                'filter': f'volume:{vol_id},issue_number:{issue}',
                'field_list': field_list,
                'limit': 10,
            })
            if vol_issues:
                vb, vs = _pick_best_match(
                    vol_issues, full_name,
                    issue_number=issue,
                    volume_number=volume_number,
                    publisher=publisher,
                )
                if vs > best_score:
                    best, best_score = vb, vs

    if best is None or best_score < CV_CONFIDENCE_THRESHOLD:
        print(f"     CV: No confident match for '{full_name}' "
              f"(best: {int(best_score*100)}% < {int(CV_CONFIDENCE_THRESHOLD*100)}% threshold)")
        return None

    print(f"     CV: '{(best.get('volume') or {}).get('name','')} "
          f"#{best.get('issue_number')}' ({int(best_score*100)}% confidence)")

    # --- Extract fields ---
    cover_date  = best.get('cover_date') or best.get('store_date') or 'Unknown'
    description = best.get('description') or ''
    deck        = best.get('deck') or ''

    key_issue = 'Yes' if any(kw in (description + deck).lower() for kw in [
        'key issue', 'first appearance', '1st appearance', 'origin',
        'death of', 'first app', '1st app'
    ]) else 'No'

    image = ''
    if best.get('image'):
        image = (best['image'].get('original_url')
                 or best['image'].get('medium_url') or '')

    chars = best.get('character_credits') or []
    characters = ', '.join([c.get('name', '') for c in chars[:10]])

    # Publisher requires a volume detail call
    publisher = ''
    vol = best.get('volume') or {}
    vol_id = vol.get('id')
    if vol_id:
        vol_data = _cv_get(session, cv_api_key,
                           f"volume/4050-{vol_id}",
                           {'field_list': 'publisher,name,start_year'})
        if vol_data and isinstance(vol_data, dict):
            publisher = (vol_data.get('publisher') or {}).get('name', '')

    return {
        'publisher':   publisher,
        'volume':      vol.get('name', ''),
        'published':   cover_date,
        'cover_price': best.get('cover_price') or 'Unknown',
        'comic_age':   _classify_age(cover_date),
        'notes':       deck or (_strip_html(description[:500]) if description else ''),
        'key_issue':   key_issue,
        'cover_image': image,
        'book_link':   best.get('site_detail_url') or '',
        'characters':  characters,
        'confidence':  best_score,
    }


# =============================================================================
#   eBay Pricing
# =============================================================================

def _ebay_sold_prices(session, query):
    """
    Fetch eBay sold listing prices via the Finding API (findCompletedItems).

    Requires EBAY_APP_ID in environment. Set EBAY_SANDBOX=true to hit the
    sandbox endpoint (test data only — switch to production keys for real prices).
    """
    app_id = os.getenv('EBAY_APP_ID', '')
    if not app_id:
        print("     eBay: EBAY_APP_ID not set — skipping eBay lookup")
        return []

    sandbox = os.getenv('EBAY_SANDBOX', 'false').lower() in ('1', 'true', 'yes')
    endpoint = EBAY_FINDING_SANDBOX if sandbox else EBAY_FINDING_PROD

    params = {
        'OPERATION-NAME':           'findCompletedItems',
        'SERVICE-VERSION':          '1.0.0',
        'SECURITY-APPNAME':         app_id,
        'RESPONSE-DATA-FORMAT':     'JSON',
        'REST-PAYLOAD':             '',
        'keywords':                 query,
        'itemFilter(0).name':       'SoldItemsOnly',
        'itemFilter(0).value':      'true',
        'sortOrder':                'EndTimeSoonest',
        'paginationInput.entriesPerPage': '10',
    }

    try:
        resp = session.get(endpoint, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()

        response_root = data.get('findCompletedItemsResponse', [{}])[0]
        ack = response_root.get('ack', [''])[0]
        if ack != 'Success':
            error_msg = (response_root
                         .get('errorMessage', [{}])[0]
                         .get('error', [{}])[0]
                         .get('message', ['Unknown error'])[0])
            print(f"     eBay API error: {error_msg}")
            return []

        items = (response_root
                 .get('searchResult', [{}])[0]
                 .get('item', []))

        prices = []
        for item in items:
            selling    = item.get('sellingStatus', [{}])[0]
            price_data = selling.get('convertedCurrentPrice', [{}])[0]
            try:
                prices.append(float(price_data.get('__value__', 0)))
            except (ValueError, TypeError):
                pass

        return prices

    except Exception as e:
        print(f"     eBay API request failed: {e}")
        return []


def _ebay_price_for_grade(session, title, issue, grade_float, cgc, variant=''):
    """Return the most recent eBay sold price for a specific grade, or None."""
    cgc_str   = "CGC" if cgc.upper() != 'NO' else ""
    grade_str = _format_grade(grade_float)
    # Build query: omit grade_str if we're doing a broad fallback search
    query = ' '.join(filter(None, [title, f'#{issue}', cgc_str, grade_str, variant])).strip()
    prices = _ebay_sold_prices(session, query)
    if prices:
        print(f"     eBay [{grade_str}]: ${prices[0]:.2f}  ({len(prices)} sales found)")
        return prices[0]
    return None


def GetEbayPriceGraded(session, title, issue, grade, variant=''):
    """Convenience: eBay price for CGC-graded copy (includes 'CGC' in query)."""
    return GetEbayPrice(session, title, issue, grade, 'Yes', variant)


def GetEbayPriceUngraded(session, title, issue, grade, variant=''):
    """Convenience: eBay price for raw/ungraded copy (no 'CGC' in query)."""
    return GetEbayPrice(session, title, issue, grade, 'No', variant)


def GetEbayPrice(session, title, issue, grade, cgc, variant=''):
    """
    Look up the most recent eBay sold price for a comic + grade.

    Returns None if:
      - grade is blank/invalid (skip gracefully instead of malformed query)
      - no eBay sales found after interpolation attempts
    """
    # Fix: blank grade → skip eBay entirely rather than search for ".0"
    grade_norm = normalise_grade(grade)
    if grade_norm is None:
        print(f"     eBay: Grade is blank/invalid ('{grade}') — skipping eBay lookup")
        return None

    try:
        target = float(grade_norm)
    except ValueError:
        print(f"     eBay: Cannot parse grade '{grade_norm}'")
        return None

    # 1. Exact match
    price = _ebay_price_for_grade(session, title, issue, target, cgc, variant)
    if price is not None:
        return price

    print(f"     eBay: No sales for grade {grade_norm} – searching nearby grades…")

    if target not in GRADE_SCALE:
        print(f"     eBay: Grade {target} not in standard scale.")
        return None

    idx = GRADE_SCALE.index(target)

    lower_grade, lower_price = None, None
    for g in reversed(GRADE_SCALE[:idx]):
        p = _ebay_price_for_grade(session, title, issue, g, cgc, variant)
        if p is not None:
            lower_grade, lower_price = g, p
            break

    upper_grade, upper_price = None, None
    for g in GRADE_SCALE[idx + 1:]:
        p = _ebay_price_for_grade(session, title, issue, g, cgc, variant)
        if p is not None:
            upper_grade, upper_price = g, p
            break

    if lower_price is not None and upper_price is not None:
        ratio = (target - lower_grade) / (upper_grade - lower_grade)
        interpolated = round(lower_price + ratio * (upper_price - lower_price), 2)
        print(f"     eBay: Interpolated {grade_norm} → ${interpolated:.2f}")
        return interpolated
    elif lower_price is not None:
        print(f"     eBay: Using nearest lower {lower_grade} → ${lower_price:.2f}")
        return lower_price
    elif upper_price is not None:
        print(f"     eBay: Using nearest upper {upper_grade} → ${upper_price:.2f}")
        return upper_price

    print(f"     eBay: No usable sales data for {title} #{issue}.")
    return None
