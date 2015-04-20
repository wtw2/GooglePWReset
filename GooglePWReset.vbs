''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Versions
'
' Version: 1.8.1
' source: inzi.com
'
' Change Log
'
' v1.0 
' Original
'
' v1.5  10/23/2013
' Added iSlowConnectionFactor, an easy way to tweak speed of script
' Updated default path to Chrome
' Added 2nd tab to account for lost password link at line 109
'
'
' v1.7 11/18/2014
' Changed path to Chrome.exe to just be "chrome.exe" Full path seems to be an issue for some users
' Increased delay for google login
' apparently - the google login is fluxing - added commenting so people can modify the dialog navigation. 
' Fixed the final password reset step so it tabs twice
' Added detailed of comments so people know what it's doing
' 
'
' v1.8 11/18/2014
' change the login URL so that the email field is editable on each login
' 
' v1.8.1 11/18/2014
' Bug - left out a quote
'
' v2.0 04/18/2015
' major update by
' Walk the Way

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' When launching this, you might pipe the output to a file in case your computer turns off,
'   so you know what the password is
' Do not activate another window during this process.  It requires continuous focus.
'

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Dependencies:  AutoItX3.dll
'
'  04/15/2015 Version 3.3.12.0 (June 1, 2014 release)
'  https://www.autoitscript.com/site/autoit/downloads/
'  1) Download ZIP
'  2) Unzip the file
'  3) For 32 bit system use:
'  /install/AutoItX/AutoItX3.dll
'  3) For 64 bit system use: (only the file name changes)
'  /install/AutoItX/AutoItX3_x64.dll
'  4) Copy to: C:\autoit\
'  5) Run CmdShell as Admin
'  6) >cd C:\Windows\system32
'  7) >regsvr32 C:\autoit\AutoItX3_x64.dll
'  Done.

'  To run in debug:
'  Open CmdShell in location of script (Shift Right Click > Open command window here)
'  >cscript /x /d GooglePWReset.vbs

'  To run regularly:
'  Open CmdShell in location of script
'  >cscript GooglePWReset.vbs

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declare Variables & Objects
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim oShell
Dim oAutoIt

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialise Variables & Objects
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set oShell = WScript.CreateObject("WScript.Shell")
Set oAutoIt = WScript.CreateObject("AutoItX3.Control")

WScript.Echo "This script will reset your Google password 100 times so you can use an old password."

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You should only edit value after this
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Your Google username (email address)
sUN = "******@gmail.com" 

' Replace with the Current password to your google account
curPW = "######" 

' Replace with the final password you want assigned to the account. The one you want to set the account's password *back to*. 
oldPW = "######" 

' Is it going to fast? You can slow it down by adjusting this value.
' If you set it to 2, it will run twice as slow
' So if it is entering data into the wrong fields, try increasing this.
' It might help.
' iSlowConnectionFactor = 1
iSlowConFac = 1

' If your password has a quote in it ("), then use "" in its place.
' For example, let's say your password was 
' MyPass"word!-55
'
' The proper VBScript way to put that into a variable would look like this
' curPW = "MyPass""word!-55"
'
' See Microsoft's website for more detail

' Where is the Chrome executable? Replace this with its location.
'     Point app to Chrome Manually
'     An easy way to find this is to right click the Chrome shortcut and copy the value in Target.
'     Click Start, type Chrome, right click Google Chrome, click Properties, copy *everything* in Target, and put it here.

' This example path is for 64 bit windows
' You might need the full path: ChromeEXE = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
ChromeEXE = "chrome.exe"


' This example path is for 32 bit windows
' "C:\Program Files\Google\Chrome\Application\chrome.exe" 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You should not have to edit anything after this
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Start of Script
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Some of this code uses the AutoIT com object. See their documentation for more details.
' https://www.autoitscript.com/site/autoit/
' you have to install that to use this.

sUrl = " --new-window -url https://accounts.google.com/ServiceLogin?Email="
iReturn = oShell.Run(ChromeEXE & sUrl & sUN, 0, True)
If iReturn <> 0 Then
    WScript.Quit
End If

' Wait for the Google Chrome window to become active
Dim sApp(5)
'sApp(0) = "New Tab - Google Chrome"					' may not last long (no url specified)
'sApp(0) = "Google - Google Chrome"						' 
sApp(0) = "Apps - Google Chrome"						' this is my default
sApp(1) = "Sign in - Google Accounts - Google Chrome"	' go straight to login page
sApp(2) = "Account settings - Google Chrome"			' page following Sign In
sApp(3) = "Password - Google Chrome"					' page following Sign In
' this is also useful to try to keep the focus on the Chrome window
'   in case a window pops up and takes focus
sApp(4) = "DEBUG" 										' for debug
oAutoIt.WinWaitActive sApp(1), "", 10 * iFac			'open chrome and wait for the window to appear
hWndw = oAutoIt.WinGetHandle(sApp(1),"")				' for debug OR try to keep in focus
result = WinVerify(oAutoIt, hWndw, iFac, sApp, 1)		' verify right window is active

WScript.Echo "Entering Loop"
tCurPw = curPW
If iSlowConFac < 1 Then iSlowConFac = 1

' Enter the loop, change the password 99 more times
WScript.Echo "Current PW: " & tCurPw
For x = 1 To 99
    WScript.Echo "Step " & x
    tNewPW = curPW & x
    WScript.Echo "Setting the password to: " & tNewPW	' output the password in case it crashes the user can see what it is.

    GLogin oAutoIt, sUN, tCurPw, hWndw, iSlowConFac, sApp, x	' log into google with current password				
    GChangePW oAutoIt, hWndw, iSlowConFac, sApp					' load the change password page, old password should have focus after it loads.
    GEditPW oAutoIt, tCurPw, tNewPw, hWndw, iSlowConFac, sApp	' change the password
    GLogout	oAutoIt, hWndw, iSlowConFac, sApp					' logs out of google - it has to do this for password change to stick.
    WScript.Echo "New PW: " & tNewPw

    tCurPw = tNewPW										' updates the current password.
Next 													' At this point, the process should repeat.

WScript.Echo "Final Change"									' Last time
GLogin oAutoIt, sUN, tCurPw, hWndw, iSlowConFac, sApp, x	' log into google with current password.
GChangePW oAutoIt, hWndw, iSlowConFac, sApp					' load the change password page
GEditPW oAutoIt, tCurPw, oldPW, hWndw, iSlowConFac, sApp	' change the password
GLogout	oAutoIt, hWndw, iSlowConFac, sApp					' logs out of google - it has to do this for password change to stick.

WScript.Echo "Password reset"

WScript.Quit

' Script complete.

Function WinVerify(oAutoIt, hWndw, iFac, sApp, x)
    list = oAutoIt.WinList("[REGEXPTITLE:(?i)(.*chrome.*)]")	' find all Chrome windows
    For i = 1 To UBound(list, 2)
        If Mid(hWndw, 3) = list(1,i) Then
            ' WinActivate doesn't work with handles; this is dumb
            WActivate = list(0,i)
            If sApp(UBound(sApp)-1) = "DEBUG" Then
                oAutoIt.WinActivate(list(0,i)) ' for debug
            End If
            oAutoIt.Sleep 250 * iFac
            Exit For
        End If
    Next
    If StrComp(WActivate, sApp(x), 1) <> 0 Then
        WScript.Echo "Exiting looking for: " & sApp(x)
        WScript.Quit		' Abort, something didn't work
    End If
End Function

Function GLogin(oAutoIt, un, pw, hWndw, iFac, sApp, i)	' Opens the Google Login page, enters the supplied Username (un) and Password (pw), and presses Enter.
    WScript.Echo "Logging in: " & un & "; " & pw
    result = WinVerify(oAutoIt, hWndw, iFac, sApp, 1)
    'If i = 1 then ' already at the login page (see above) 
    '    oAutoIt.Send "!d"			' This goes to the address bar
    '    'oAutoIt.Send "https://accounts.google.com/Login{ENTER}"		' types this url and hits enter. Upon load, email field should have focus
    '    oAutoIt.Send "https://accounts.google.com/ServiceLogin?Email=" & un & "{ENTER}" ' types this url and hits enter. Upon load, email field should have focus. Email param makes it empty.
    '    oAutoIt.Sleep 3000 * iFac 	' waits x ms times slow connection
    'End If
    oAutoIt.Send pw & "{ENTER}"		' types password and hits enter
	oAutoIt.WinWaitActive sApp(2), "", 10 * iFac 
    oAutoIt.Sleep 500 * (iFac-1)	' wait for the page to render
End Function

Function GChangePW(oAutoIt, hWndw, iFac, sApp) 				' Opens the Google Change Password web page
    WScript.Echo "Requesting password change."
    result = WinVerify(oAutoIt, hWndw, iFac, sApp, 2)
    oAutoIt.Send "!d"				' go to address bar
    oAutoIt.Sleep 200 * iFac 		' waits x ms times slow connection
    oAutoIt.Send "https://accounts.google.com/b/0/EditPasswd{ENTER}"	' go the edit password page
	oAutoIt.WinWaitActive sApp(1), "", 10 * iFac 
    oAutoIt.Sleep 500 * (iFac-1)	' wait for the page to render
End Function

Function GEditPW(oAutoIt, pw1, pw2, hWndw, iFac, sApp) ' Opens the Google Change Password web page
    WScript.Echo "Changing password."
    result = WinVerify(oAutoIt, hWndw, iFac, sApp, 1)
    oAutoIt.Send pw1 & "{ENTER}"	' enter the old (last) password (required confirmation login)
	oAutoIt.WinWaitActive sApp(3), "", 10 * iFac 
    oAutoIt.Sleep 500 * (iFac-1)	' wait for the page to render
    result = WinVerify(oAutoIt, hWndw, iFac, sApp, 3)
    oAutoIt.Send pw2				' type new password and tab THRICE
    oAutoIt.Sleep 100 * iFac 		' waits x ms times slow connection
    oAutoIt.Send "{TAB}"			' type new password
    oAutoIt.Sleep 100 * iFac 		' waits x ms times slow connection
    oAutoIt.Send "{TAB}"			' type new password
    oAutoIt.Sleep 100 * iFac 		' waits x ms times slow connection
    oAutoIt.Send "{TAB}"			' type new password
    oAutoIt.Sleep 100 * iFac 		' waits x ms times slow connection
    oAutoIt.Send pw2				' types the new password again then tab TWICE
    oAutoIt.Sleep 100 * iFac 		' waits x ms times slow connection
    oAutoIt.Send "{TAB}"			' type new password
    oAutoIt.Sleep 100 * iFac 		' waits x ms times slow connection
    oAutoIt.Send "{TAB}"			' type new password
    oAutoIt.Sleep 100 * iFac 		' waits x ms times slow connection
    oAutoIt.Send "{ENTER}"			' Hit enter to submit the password reset.
	oAutoIt.WinWaitActive sApp(2), "", 10 * iFac 
    oAutoIt.Sleep 500 * (iFac-1)	' wait for the page to render
End Function

Function GLogout(oAutoIt, hWndw, iFac, sApp) ' Logs out from google. This is necessary for the password change to take effect.
    WScript.Echo "Logging out."
    result = WinVerify(oAutoIt, hWndw, iFac, sApp, 2)
    oAutoIt.Send "!d"				' this goes to the address bar
    oAutoIt.Sleep 200 * iFac 		' waits x ms times slow connection
    oAutoIt.Send "https://www.google.com/accounts/Logout{ENTER}"	' This logs out of google. 
	oAutoIt.WinWaitActive sApp(1), "", 10 * iFac 
    oAutoIt.Sleep 500 * (iFac-1)	' wait for the page to render
End Function