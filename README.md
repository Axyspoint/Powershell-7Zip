# Powershell-7Zip
Script is compatible with Tanium and other 32-bit Clients.
Uninstall EVERY Version of 7-Zip (As long as it exists in the uninstall registry). 
It will then install 7-Zip 1805. 

The version is manually put into the script. (Lines: 8,29,34,37,45,54,58,66). 

You could put the version into a variable at the beginning of the script if you wish and replace any 1805 with the variable.

Also note, that the script might (Unlikely) fail to uninstall if the uninstall string has quotes in it.
I've seen a few that didn't uninstall due to quotes, however, with what I've tested, I've had a great success on 30 different computers with different quoted Uninstall strings.
