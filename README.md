# Data-Analysis-Projects

To help new or novice users have a step-by-step tutorial on how to create and run a macro in Microsoft Excel that builds an inventory of files found within a folder File name, file type, author, creation date time, modification date-time, size of the files and path will be gathered & rendered in an organized spreadsheet as shown above on functionancial. Subsequent documentation will make it possible for users to turn on grid features as required run the macro and troubleshoot simple issues.

-for New Users follow the steps in enabling the "Developer tab"
-for Beginners follow all this steps.
-skip to no. 8 if it only needs to refresh

#### Prerequisites
- Microsoft Excel (preferably version 2010 or later)
- Basic knowledge of how to open and edit Excel files
- The folder containing the files to be cataloged

#### Steps to Set Up and Run the Macro

1. **Open Excel:**
   - Launch Microsoft Excel.

2. **Enable the Developer Tab:**
   - If the Developer tab is not visible, enable it by going to `File > Options > Customize Ribbon`.
   - Check the `Developer` option in the right pane and click `OK`.

3. **Open the VBA Editor:**
   - Click on the `Developer` tab.
   - Click on `Visual Basic` to open the VBA editor.

4. **Insert a New Module:**
   - In the VBA editor, go to `Insert > Module` to create a new module.

5. **Paste the VBA Code:**
   - Copy the VBA code provided below and paste it into the new module.

6. Save the Excel file.

7. Close and reopen the Excel file, enabling macros if prompted.

8. Run the script as before:
   - Go to the "Developer" tab
   - Click "Macros"
   - Select "GetFileInfoRecursive" and click "Run"

9. Select the main folder you want to catalog when prompted.

10. Wait for the script to complete. It will show a message when done.

11. Review the results in your spreadsheet. The columns should now be in this order:
   A: File Name
   B: File Type
   C: Author
   D: Date Created
   E: Date Modified
   F: File Size
   G: File Location (as a hyperlink)


#### Troubleshooting
- Ensure the Developer tab is enabled and visible.
- Make sure you have selected a folder when prompted by the file dialog.
- Check that the files in the selected folder have one of the supported extensions.
- If you encounter any errors, re-check the VBA code for any missing or incorrect lines.
