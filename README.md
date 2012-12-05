SAP Mass Password Change Script
===============================

Based on <http://ceronio.net/2007/08/mass-change-your-sap-passwords/>

This is a VB script that allows you to mass change your SAP passwords by logging on to each system and calling
SUSR_USER_CHANGE_PASSWORD_RFC using RFC.

The requirement is that before you begin, all your passwords must be the same on all systems.

The script prompts you for the old password, then the new password, and then proceeds to log into each system you have specified
and change the password to the new password given.

A log file is written showing the outcome of each pasword change.

## Usage

1. Make sure all your passwords are the same on all systems to begin with!

2. Change the script where indicated to add a line for each system on which you want to change your password.
   The fields are: host, system ID, system number, client, language, user ID

3. Run the script from the command line, supplying the old password and a new password

After running the script, check the log file 

## Windows 64 bit notes

(Thanks to Mark Bailey for this information):

Use the included batch file on 64 bit Windows.