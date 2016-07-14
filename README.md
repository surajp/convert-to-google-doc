# convert-to-google-doc
Nifty nodejs script to convert office documents to Google Docs, and place a shortcut in the same folder. Built to work as a convenient context menu item in Windows.

# Points to note
* Conversion Works with .doc,.docx,.xls,.xlsx,.csv,.ppt,.pptx and .txt files. All other files are uploaded as is.
* Works best when you have a single chrome profile on your desktop. With multiple profiles, you may see some access issues, based on which profile was the last one accessed.
* The documents are uploaded to a folder by the name *GConvert* under the root folder. Gets created if it doesn't exist already.
* I did see the shortcut being created incorrectly on some machines. If that happens you may have to hardcode the full path to chrome.exe in the script.
* Install Google Docs, Google Sheets and Google Slides chrome extensions for offline access.
* Haven't tested this with really large files yet. If you want that supported, let me know.



# Steps to get the script running
1. Clone the repo and run `npm install` to install the dependencies.
2. Make sure chrome.exe is in your system path.
3. Run the command `node theScript.js "[Path to word or excel file]"`
4. Follow on-screen instructions for the first run to authenticate and authorize the application.
5. You should see your converted document open up, and a shortcut appear in the same folder.

# To add a context menu item
* Follow the instructions [here](http://www.howtogeek.com/107965/how-to-add-any-application-shortcut-to-windows-explorers-context-menu/)
* Use this as the command. Make sure you include the quotes as is: `"[full path to node.exe]" "[full path to the script]" "%1"`

