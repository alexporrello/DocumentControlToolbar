# Document Control Boilerplate Macros

## 1. Installing the Macros

### 1.1 Enable the _Developer_ Tab

To use the macros in this repository, you must enable the _Developer_ tab in Microsoft Word:

1. Naviage to _File_ >> _Options_ to open the _Word Options_ screen.
2. Navigate to the _Customize Ribbon_ window.
3. In the _Main Tabs_ list on the right side, check the box next to _Developer_.
4. To save your changes and exit the window, click the _OK_ button.

![](readme_screenshots/customize-ribbon.PNG)

Now, in the toolbar in Microsoft Word, you should see a _Developer_ tab:

![](readme_screenshots/developer-tab.PNG)

### 1.2. Download the Files

1. Visit the _release_ tab on the GitHub repository page.
![](readme_screenshots/releases.PNG)
2. Underneath _Assets_, click the _Source code (zip)_ link to download the macros.
![](readme_screenshots/release-page.PNG)
**Note:** You can confirm that you are downloading the latest version by locating the green box that says _Latest release_.
3. Navigate to your default download location on your computer. The downloaded file should be called _TWBoilerplateMacros.zip_.
4. Extract the files from the downloaded _.zip_ folder.

### 1.3. Import the Files into Microsoft Word

1. Open Microsoft Word.
2. Navigate to the _Developer_ tab.
3. Click on _Visual Basic_ button to the far left of the ribbon.
![](readme_screenshots/visual-basic-button.PNG)
This action opens Visual Basic for Applications (VBA):
![](readme_screenshots/word-vb.PNG)
4. In VBA, navigate to _File_ >> _Import File..._
5. Browse to the directory where you unzipped the macros (most likely your _Downloads_ directory):
![](readme_screenshots/import-file.PNG)
6. If you are running for the first time, import all of the visible files in _forms_ and all of the items in _modules_. Since VBA does not support bulk importing, you will have to import them individually.

Assuming you followed the above steps correctly, all of the macros should have been imported into your normal template.