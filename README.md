# Excel Encryption Application (Crazy 'Cryption)

## Description:

In my application called Crazy 'Cryption, you are given three buttons to use for creating a new worksheet, encrypting the data in those created worksheets, and finally, decrypting those worksheets and creating a word doc with a summary of the database results and a plot chart displaying the encryption stats.

The program encrypts the worksheets into a .crzy file that is unreadable. When the user clicks the decrypt button, the application goes through the files and decrypts them into a .notcrzy file which is readable with all the data stored in the encrypted .crzy file.

While encrypting the worksheets, the application has a counter for various stats like the files encrypted and how much data has been encrypted – to be used for the word doc plot chart. The counter stats are sent to the local Access database which some SQL commands to send and receive the data.

## Button Info:

### Create New Worksheet Button (plus button):

- Used to create a new blank worksheet where data is entered that will be encrypted
- Can have data copy-pasted for already existing worksheets you would like to encrypt.

### Encryption Button (locked button):

- Used to encrypt all current (active) worksheets to a .crzy (unreadable) file
- The only way to read the encrypted file is to decrypt its contents by using the decryption button (See Decryption Button feature)

### Decryption Button (unlocked button):

- Used to decrypt .crzy files in the application folder (where the excel file is) to a readable .notcrzy worksheet
- When clicked, a word doc is created with a summary of the encryption stats (Files encrypted counter, data encrypted counter, worksheet name, and worksheet encryptions – encryptions/worksheet) collected during the encryption phase and displays a plot chart

## How to Use:

1. Before we start, please ensure you have enabled macros in the file for everything to run.

2. First, start by creating a new worksheet (clicking the + button). This worksheet can have whatever data you wish to encrypt. To encrypt an already existing worksheet, simply copy and paste its contents into the newly created worksheet.

3. Now click the lock button to encrypt the data on all active worksheets. You will notice .crzy files with their respective worksheet names in your file explorer, however, they cannot be opened.

4. Lastly, click the unlocked button to decrypt the .crzy files to a readable .notcrzy worksheets and the plot chart worksheet from the database. These worksheets will display the contents of the .crzy.

5. You will notice a word document appear with a plot chart and table summarizing the encryption stats (see decryption button feature for more info or below for a screenshot of an example doc).

## Known Issues:

- You may run into an issue where an error message will pop up if you decrypt the same files multiple times while the created word doc is open or saved with the same name as a new one. To avoid this, please close or save the word doc with another name before clicking decrypt again as well as deleting the plot chart worksheet (to keep its contents, please export it with a different name then proceed to delete it)

## What I Learned:

As a second-year CS student, I am exploring different fields to specialize in. I’ve always been interested in the field of cybersecurity. That’s what lead me to the Crazy ‘Cryption application as it seemed super interesting to get my feet wet for how data is encrypted. Creating basic data encryption taught me the general process to encrypting data. On top of that, making it all pretty and user-friendly is also something I like to do. My goal for the project was to make the front-end as uncomplicated as possible.

This also exposed me to how using databases with a project and handling data instead of storing it in the project itself. The Access database I used has SQL commands to send and retrieve the encrypting stats used for the plot chart.
