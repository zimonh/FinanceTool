# FinanceTool

## Finance Tool - How to Install

### Important: Use Google Chrome
<img src="https://blog.zimonh.at/wp-content/uploads/2018/04/Google_Chrome-300x300.png" width="100px" height="100px">
<p>During this installation it's important you use the Chrome browser.
<br> If you do not have Chrome installed you can download it here: <a href="https://www.google.com/chrome/">www.google.com/chrome/</a>
<br> After this installation you can use any browser.</p>

### Before: Sign in
<img src="https://blog.zimonh.at/wp-content/uploads/2018/04/signin-300x197.png" width="300px" height="197px">
<p>Make sure you have a Google account and are signed in.
<br> If you don't have an account you can make one here: <a href="https://accounts.google.com/SignUp">accounts.google.com/SignUp</a></p>

### Step 1: Download
<img src="https://blog.zimonh.at/wp-content/uploads/2018/06/download.gif" width="720px" height="auto" />
<p> On this page you can find the two files: <a href="https://drive.google.com/open?id=1UlBqv72MITxK9OUbtVVfgBZX3Cq57Nzr">download</a>
<br> <b>Click</b> + <b>drag</b> and select both files,
<br> Now <b>right click</b> on one of the selected files and choose <ins>Make a copy</ins>.
<br> Wait 30 seconds for the copying to complete and the files are now part of your drive.</p>

### Step 2: Connect the Doc to the Sheet.
<img src="https://blog.zimonh.at/wp-content/uploads/2018/06/link-invoice.gif" width="720px" height="auto" />
<p>In the left menu <b>click</b> <ins>My Drive</ins>.
<br> Now <b>double click</b> <ins>Copy of Invoice</ins> to open the invoice document.
<br> Once the file is open <b>select</b> the following part of the URL: <img style="border-radius:3px" src="https://imgur.com/06cSLMu.png" width="700px">
<br> And press <b>Ctrl</b> + <b>C</b>.
<br> Now go back to <ins>My Drive</ins>.
<br> And <b>double click</b> <ins>Copy of Finance Sheet By ZIMONH</ins>.
<br> In the bottom <b>click</b> the <img src="https://blog.zimonh.at/wp-content/uploads/2018/05/tab_settings.png" alt="Settings" height=25px; width=auto> tab.
<br> Now find <ins>Google Drive file id for invoice Document:</ins>
<br> And <b>click</b> on the cell next to it and press <b>Ctrl</b> + <b>V</b>.
<br> Now refresh the page.</p>

### Step 4: Authorize
<img src="https://blog.zimonh.at/wp-content/uploads/2018/06/permission2.gif" width="720px" height="auto" />
<p>In the top menu <b>click</b> <b><ins>Tools</ins> &gt; <ins>Script editor</ins>.</b>
<br> Now in the top menu <b>click</b> <b><ins>Edit</ins> &gt; <ins>Current project's triggers</ins>.</b>
<br> Now a prompt will appear, <b>click</b> <ins>No triggers set up. Click here to add one now.</ins>.
<br> In the first list <b>click</b> <ins>createDoc</ins> and select <ins>onEdit</ins>,
<br> In the second list <b>click</b> <ins>Time-driven</ins> and select <ins>From spreadsheet</ins> and
<br> In the third list <b>click</b> <ins>On open</ins> and select <ins>On Edit</ins>.
<br> <b>Click</b> <ins>Save</ins>.
<br> Now a prompt will appear saying: Authorization required.
<br> This is needed to allow the finance sheet to connect to the invoice doc.
<br> <i>This script will never send information to 3rd parties or try to access other files.</i>
<br> <b>Click</b> <ins>Review Permissions</ins> and select your account.
<br> Now a warning will appear.
<br> In the bottom <b>click</b> <ins>Advanced</ins>.
<br> Next scroll down and <b>click</b> <ins>Go to Billy (unsafe)</ins>.
<br> Now a prompt will apear asking permission for this script: <b>Click</b> "Allow".
<br> What permissions you are granting? View and manage the files in your Google Drive.
<br> <i>This is needed to link the Document to the sheet and to add a signature image.</i>
<br> View and manage your spreadsheets in Google Drive.
<br> <i>This is needed to make changes to the finance sheet.</i>
<br> View and manage your documents in Google Drive.
<br> <i>This is needed to make changes to the Invoice document.</i>
<br> </p>

### Step 5: Test
<img src="https://blog.zimonh.at/wp-content/uploads/2018/06/test.gif" width="720px" height="auto" />
<p> <b>Click</b> <ins>X</ins> On the current <ins>Billy</ins> tab in your browser.
<br> And in the <ins>Finance sheet</ins> go to the <img src="https://blog.zimonh.at/wp-content/uploads/2018/05/tab_finance.png" alt="Finance" height=25px; width=auto> tab.
<br> Now <b>double click</b> the <ins>Bill</ins> column of row <ins>5</ins>.
<br> And <b>click</b> on a date.
<br> Wait 30 sec.
<br> Now open the <ins>Copy of Invoice</ins> and check if the date is correct.</p>

### Step 6: Add Logo
<img src="https://blog.zimonh.at/wp-content/uploads/2018/06/logo.gif" width="720px" height="auto" />
<p>Make sure you have a logo with a white or transparent background on your computer.
<br> <b>Click and Drag</b> the file into <ins>Copy of Invoice</ins> next to the logo.
<br> <b>Click</b> the old logo and press <b>Delete</b>.
<br> Now <b>Click</b> the new logo and adjust the size.
<br> The file will automatically save.</p>

### Step 7: Add Signature
<img src="https://blog.zimonh.at/wp-content/uploads/2018/06/signature.gif" width="720px" height="auto" />
<p>Make sure you have a signature with a white or transparent background on your computer.
<br> <b>Click and Drag</b> the file into <ins>My Drive</ins>.
<br> <b>Right Click</b> on the new file and choose <ins>Get sharable link</ins>.
<br> <b>Select</b> the text after <ins>id=</ins> and press <b>Ctrl</b> + <b>C</b>.
<br> Now open <ins>Copy of Finance Sheet By ZIMONH</ins> go to the <img src="https://blog.zimonh.at/wp-content/uploads/2018/05/tab_settings.png" alt="Settings" height="25px" width="auto" /> tab.
<br> And <b>Click</b> next to <ins>File id for Signature Image</ins> and press <b>Ctrl</b> + <b>C</b>.
<br> Now when you send a final notice or a bill with late fee it will have your signature. </p>

### Step 8: Add Company Info
<img src="https://blog.zimonh.at/wp-content/uploads/2018/06/settings_basic.gif" width="720px" height="auto" />
<p>In the <img src="https://blog.zimonh.at/wp-content/uploads/2018/05/tab_settings.png" alt="Settings" height=25px; width=auto> tab you can add your company info.
<br> This will all appear when you create an invoice.
<br> You can adjust all the text that doesn't have a black background.</p>

### Step 9: Other Customization
<img src="https://blog.zimonh.at/wp-content/uploads/2018/06/settings_translate.gif" width="720px" height="auto" />
<p>You can also translate, adjust the tax-rate and the currency. For more information visit <a title="How to use" href="https://blog.zimonh.at/finance-tool-how-to-use/">How to use</a>.</p>

## License
<a rel="license" href="http://creativecommons.org/licenses/by-nc-sa/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by-nc-sa/4.0/88x31.png" /></a><br>
Copyright - ZIMONH 2018
