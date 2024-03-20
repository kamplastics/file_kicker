# File Kicker

Ever have it on a windows network where there is a shared file that is open by another user on your network and you are now kinda stuck and you need to kick them out (usually excel files it seems). 
You can call the IT guy and have him log into the shared box, go to computer management and manually kick them out, but who has time for that.

The file kicker aims to make a normie friendly method to kick people out of files.

Create the directory + file  kick\kick.txt in the folder you wish to kick people out of.

Then create a scehduled task to this application that runs with the highest priviliages. I set mine to run every 5 minutes and told the end users to wait 5 minutes after setting the `kick.txt` file to 1. 

## Legal

I assume no responsibilty for this application. Read the source code yourself and use at your own risk. 
Project contributions/pull requests are welcome. 
