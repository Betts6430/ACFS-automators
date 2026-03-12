# COA and Product Key Autofiller


### Description
Upon running the script, the user is asked for a URL to the Google Sheet. The script will monitor for changes in the O:/ (Output) folder and detect new .report/.assemble files. If the serial number found in the .report file matches with one in the sheet, that device will have its New COA and Product Key autofilled. Only one person should run the script at once.

### How to use 
Copy and paste the URL of the current Google Sheet containing information on refurbished computers (for 2026, it should autofill the correct URL already)
Make sure the serial number of any computer you wish to use this script on is in the “Serial #” column before running the script
That’s it! The script will autofill all COAs and Product Keys as long as it is running.

<img width="434" height="255" alt="image" src="https://github.com/user-attachments/assets/2e2388b3-62ea-4cbe-b04f-aa10496d2059" />

### When it will fail 
- If you input an incorrect URL
- If there are duplicate serial numbers (this should not be possible, however)
- If you incorrectly type the serial number (or forget it altogether)
- If the script cannot find the serial number of the device, nothing will happen and it will move on 
- If a user has their cursor in a field that the script is trying to populate

### Note to Supervisors 
The columns: ‘Serial #’, ‘New COA’, and ‘Product Key’ must have these exact names (although they can be in any order) for the script to detect where to paste values. The script is designed to only check the current/live month for entries. If you make a new Google Sheet, you must add ‘sheets-bot@sheets-automation-485701.iam.gserviceaccount.com’ as an editor.

Script Developed by Izaan : )
