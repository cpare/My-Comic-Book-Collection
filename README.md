# My-Comic-Book-Collection
  Python script that takes just a few inputs from an Google sheet to query for additional details from https://comicspriceguide.com.  
Result is an updated sheet with current values, description, and other helpful features as well as a helpful HTML page with your entire collection in s nice visual layout.

<b>Provide the following items in an XLS:</b>
 - <b>Title</b> - Your Comics Title (Amazing Spider Man)
 - <b>Issue</b> - The Issue Number (101)
 - <b>Grade</b> - Grade of the book based on numbers offered on https://comicspriceguide.com
 - <b>CGC</b> - Yes/No if the Book has been "Officially graded"
 - <b>Variant</b> - Some Issues have multiple printes/covers, this is where yuo indicate the specific cover you have (E)
 - <b>Price_Paid</b> - The Price you paid for the book
 - <b>Book Link</b> - Sometimes the app doesn't select the right book, this allows you to specifically provide the URL, when populated we skip the search and use this (Optional)
 
 <b>XLS Results:</b>
  - <b>Publisher</b> - Publisher of the title
  - <b>Title</b> - (Provided Above)
  - <b>Volume</b> - Volume 
  - <b>Issue</b> - (Provided Above)
  - <b>Variant</b> - (Provided Above)
  - <b>Grade</b> - (Provided Above)
  - <b>CGC</b> - (Provided Above)
  - <b>PublishDate</b> - Original Publish Date
  - <b>KeyIssue</b> - Is the book considered a "Key Issue" (Yes/No)
  - <b>Price_Paid</b> - (Provided Above)
  - <b>Cover_Price</b> - Original Cover Price
  - <b>Graded Price</b> - Value of book & and grade if professionally graded
  - <b>Ungraded Price</b> - Value of book if not professionally graded
  - <b>Value</b> - Value of book in current state
  - <b>Comic_Age</b> - Comic "Age" (Silver, Modern)
  - <b>Notes</b> - Short description of the comic, including references to 1st appearances
  - <b>Confidence</b> - Confidence the app had when attempting to match
  - <b>Book Link</b> - (Provided Above)
  - <b>Cover Url</b> - Cover Image
  - <b>Date</b> - The date the scan was performed and the value of the book at that time, allowing you to scan periodically and easily identify price fluctuations.

<b>HTML Results:</b>
 - A single HTML page with the covers of each of your books, formatted to look stunning on any format
 - Hover-over any book to see book details (Title, Grade, Value, CGC Graded)
 - Click any book to go directly to that book on https://comicspriceguide.com

IMPORTANT NOTE: Though I have taken an effort to try and simulate a normal user session wtih delays there is a chance the bot could be detected.  If you run this and are detected by https://comicspriceguide.com you stand a chance of having your IP banned from that site.
