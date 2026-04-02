# Refurbished Computers Bot V4.2.4


### Description
Autofills computer information when entering new refurbished computers by reading from the Google Sheet. Once all computers are entered, the bot can optionally auto-enter them into an order.

#

### Setup
Before running the program, make sure `config.ini` is in the same folder as the script. This file contains the URLs the bot uses and can be edited with any text editor. Open it up and update the values if any URLs or IPs have changed.

> **Note:** `config.ini` contains sensitive information and should not be shared or committed to version control.

#

### How to use it
Upon running the application, a window will pop up. Enter the following information:
1. **Google Sheet URL** (will be different for each month's sheet)
2. The **row number** of the first computer you want to enter
3. The **number of computers** you want to autofill
4. Check or uncheck **"Auto-enter into order?"** depending on whether you want the bot to paste barcodes into an order afterwards

Once you click the **"Run program"** button, a status message will appear while the spreadsheet data is being fetched. After that, an automated Chrome browser will open to the ACFS database. You will be prompted to login and navigate to the refurbished computers page.

When you're ready, click **"Add a new computer"** and the bot will autofill the form. After you confirm the information looks good, click **"Submit"** and the bot will move on to the next computer.

If **auto-enter into order** is enabled, the bot will prompt you to select an order after all computers have been entered, then paste each barcode automatically.

After everything is complete, the program will shut down on its own.

#

### What not to do
Note that if the bot ever malfunctions, it will shut down or stop working. In both cases you can just close everything and start it back up again.

- Ensure the information on the spreadsheet is all filled in and the row number is correct. The bot will show an error if any cells are missing.
- The program will time out after 5 minutes, so don't take too long on the login page.
- Try not to stray too far in the automated browser.
- Closing the browser window manually will cancel the automation.

#

### Changing URLs and IPs
Open `config.ini` in any text editor (Notepad works fine). Each setting has a comment above it explaining what it's for. Change the value after the `=` sign, save the file, and restart the bot.

#

Script by Avery