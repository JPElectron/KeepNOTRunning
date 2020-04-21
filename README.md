# KeepNOTRunning
Run something when the computer is idle

Let's say someone leaves Outlook running, thereby chewing up bandwidth by checking their POP account for new spam every 2 minutes, when the machine is idle (they haven't moved the mouse or pressed a key in awhile) then kill off outlook.exe

Let's say someone launches NBC video in their browser, then goes to lunch, when the machine is idle kill off iexplore.exe

Let's say someone logs onto a computer as administrator, then gets distracted and leaves, when the machine is idle logoff the user.

Tested on Windows 2000, XP, Server 2003, Vista, Server 2008, Windows 7

A scripted install/uninstall is not included with this software.

This program runs in the background; without any GUI, taskbar, or system tray icon.

Since this program is 32-bit it can only detect other 32-bit applications.

For the opposite of Keep NOT Running see Keep Running

<b>Installation:</b>

1) Ensure this prerequisite is installed: Microsoft Visual Basic 6.0 SP6 Run-time Components
2) Extract the contents of the .zip file
3) Modify keepnotrun.ini as indicated below
4) Run keepnotrun.exe

<b>.ini Settings:</b>

Do not use quotes around the full path, even if it contains spaces.
Example: Detect=C:\Program Files\Microsoft Office\OFFICE11\outlook.exe

Launch= will be run if the detected program is still found running after the Idle= time.
This could be set to the path of another .exe, a .bat file which could use pskill.exe to "End Task" on the detected program, or use logoff.exe to logoff the user who obviously isn't there anymore.

<b>Usage:</b>

While Keep NOT Running works independently of any screen saver setting, it uses the same method Windows does to determine when the machine is idle and thereby launch a screen saver. If an application prevents the screen saver from starting (one example is Slingbox's SlingPlayer) then it's likely that Keep NOT Running also won't be able to detect when the machine is idle, and therefore Launch= may never run.

Example to turn off a PC when idle:  http://jpelectron.com/sample/JPElectron/KeepNotRun,%20idle%20shutdown.zip
