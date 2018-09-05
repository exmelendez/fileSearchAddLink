# File Search + Link Add: A Google Sheet Apps Script Project
**Setup:** I work at a school where I provide tech/admin support.

**Problem:** We receive students completed exams as image files along with a spreadsheet that details the students information and an ID number that corresponds to each image file. I was tasked with locating the image files on Google Drive, finding the correctly associated image, getting its shareable hyperlink and then pasting it on the spreadsheet. After executing the process individually for a few of them, I quickly realized how tedious the process is.

**Solution:** Programmatically search Google Drive, used the individual ID from each row to determine if the file is there; if it is, get it's link and automatically insert the URL into the spreadsheet so those viewing it can view the individual's students exam.
