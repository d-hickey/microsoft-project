
These files were designed to work for Microsoft Office Word 2013 which currently only works on Windows machines

Instructions for getting the application working with Word can be found here: http://msdn.microsoft.com/library/office/fp142255%28v=office.15%29
I also made some less formal instructions for the group here, step 1 is what you'd need to be looking at: http://dar.netsoc.ie/Microsoft/home.html

For the above two the file path in the XML file should link to landing.html

If you do not have access to Word 2013 or don't want to go through the trouble of setting it up you can always just run it in a web browser
as an office app is basically a web page.
It should be noted that as the app was designed for Word and not web browsers the design may be a little off, mainly the text box that the code is pasted into.
All functionality will still be correct though.

To do this you'll want to click on CodeFormatting.html which should load in your browser of choice
However you will have to change a line in the program.js.
Uncomment the statement at line 95: "document.getElementById("results").innerHTML = code;"
Also comment out the statement on line 96 below it: "//Office.context.document.setSelectedDataAsync(code, { coercionType: 'html' });"
These are the lines responsible for outputting the code to either the html page or to the Word document.

You can also use http://dar.netsoc.ie/Microsoft/Formatter/CodeFormatting.html where the javascript file has been changed appropriately.
The above address cannot be used as the Catalog URL in the Trust Center as Office will only allow https addresses.

Landing.html is set up to detect the language of the Word document and move to the appropriate start page and should be ignored if not using Word.