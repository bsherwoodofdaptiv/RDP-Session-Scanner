RDP Session scanner runs as either an automated task, or can be run manually.

This tool connects to active directory to get a list of registered Domain servers. With this list it tries to 'ping' each one. Those that do not respond are added to a list of servers that are potentially non-existe
nt (retired?) and Active Directory should be cleaned up to remove them.

For each server that does respond, a WMI connection is made and a list of all users who have an EXPLORER.EXE session running in the process list is created. These are added to a report that will get mailed out after all of the servers are processed.

This tool MUST be run as an account with high enough active directory privileges to make the WMI connection to each server, as well as poll the Active Directory domain controller for the server list. Typically this is an account with Domain Admin rights.

----------------------------------------

I originally wrote this utility out of frustration with my operations team mates. For years we would fight over abandoned RDP sessions on servers. The impact of this is that on a generic out of box Windows server, you are allowed only 2 administrative RDP session at a time to be logged onto the box. Attempts to connect to the box and log in are met with a 'too many sessions already logged in'.

With the advent of how Windows 2008 deals with this, and how you can kick another user out of the box, life is a little easier to deal with. Especially when you are the on-call rotation person for the day, and a problem crops up on that specific server.

----------------------------------------

This tool is written in Windows Scripting / Visual Basic Scripting. What it can do did not warrant writing it in C# or any other compiled language.

It leverages WMI as well.

Yes, it could probably be written in very few lines inside of PowerShell. Go for it.