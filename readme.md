Notify.vbs
==========

Description
-----------
This small VBScript sends an email from your Gmail account when executed. It's primary use is to notify you whenever your computer is logged on, and display a short message on screen to the user. 

Attributes
----------
The code has been modified from [CyberNet News](http://cybernetnews.com/vbscript-send-emails-using-gmail/) so that command line arguments are not required. Every variable is set by the user within the script file.

Configuration
-------------
Be sure to enter your details in the Configuration section at the top of the script. Details include
- Name
- Sender's Gmail address
- Sender's Gmail password
- Recipient's email address
- Email subject
- Email body text
- Message dialog box title
- Message body text

Installation
------------
1. Download the script. Make sure the saved file has a `.vbs` file extention.
2. Configure the script with a text editor as per the Configuration instructions above.
3. Place the script in a high-level location, e.g. `C:\scripts\`. This means the script can run for all users.
4. Right-click and create a shortcut.
5. Open the system Startup folder, located (for Windows 7) at `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup`.
6. Place the shortcut in the Startup folder.

Disclaimer
----------
*Your Gmail password is stored as plain text in this script.* Anyone can view this if they view the source code. Use at your own risk.

Note that this script is only compatible with computers running a Windows operating system, and has only been tested on Windows 7.

