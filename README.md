# pr2csv
Podcast Republic database to CSV converter

**Initial pupose**  
  Extract from the exported database file of Android application Podcast Republic the episode names and publication date of the podcasts.
  It will store that in a .CSV file allowing to know the correspondance of each downloaded MP3 file.

**Building**
  Application has been made using Lazarus/Free Pascal.

**Justification**
  Podcast Republic download all the MP3 files to same folder but with a kind of unique filename not chosen to be intuitive.
  User does not have idea just at looking at a filename like "40858293799.mp3" to which podcast it coresspond.
  The application PR2CSV will extract information from the Podcast Republic to give us this information.
  Thanks to Podcast Republic to use the known SQLite database structure.

**Additional functionalities**
  The .CSV file was initially used as a data to feed script to rename the MP3 we wanted the keep and having a significant filename.
  PR2CSV has been modified to incorporate a table with the infomation extracted.
  Then, application has been modified to incorporate a way to rename these file with the user's wanted filename style.
  Use will type a formula that could incorporate static text, extracted episode name and publication date.
