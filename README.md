## Folder-Organisation-by-IDs-in-Filenames-with-Excel-import

Just a small automation script for my sister, to find files with matching IDs.

![image](https://github.com/user-attachments/assets/ff3db06e-0588-442a-bcd9-bc2f81c5639e)

You select the source folder with the to be selected files, load the IDs from an .xlsx or .xls file (or type them in with separating commas manually) and select a destination folder. The source folder will be searched recursively through all subfolders.

Then it'll match the IDs to the filenames and copy the matches into the destination folder. Additionally it will generate an .xlsx report telling you the matches files, the unmatched IDs and the unmatched files.

How the matching is done here:
- IDs are integers of possibly different length. The Filename can have letters after the ID. Thus the exact ID is machted with optional letters before and after (see the regex). .png is prefered and will be matched in priority over .NEF.
- Only one match is selected for each ID

To install yourself, simply install the packages and run the .py program. Alternatively use pyinstall to create an executible for yourself. If for whatever reason you want to use this highly specific script without changing the matching and are using windows, you can contact me and I'll send you the executable, then there's no python needed on your side.
