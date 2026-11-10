-- MS Report.app: installs the tools once, then opens the MS Report window (dashboard.py).
-- All the work is done by msreport.sh inside the app; this file only shows the install dialogs.
-- A folder dropped on the app icon opens in the window.

property appTitle : "MS Report"

on run
	startApp("")
end run

on open theItems
	startApp(POSIX path of (item 1 of theItems))
end open

on startApp(target)
	set helper to quoted form of (POSIX path of (path to resource "msreport.sh"))

	-- 1. One-time install (and again after an update that needs new packages) ----------
	if (do shell script "/bin/bash " & helper & " ready") is not "ready" then
		set firstTime to (do shell script "/bin/bash " & helper & " installed_before") is "no"
		if firstTime then
			display dialog "Welcome to MS Report." & return & return & ¬
				"Before the first use, MS Report downloads the tools it needs (about 600 MB, internet needed). " & ¬
				"This takes 2-5 minutes and happens only once. A progress window shows where it is." ¬
				buttons {"Cancel", "Install"} default button "Install" with title appTitle with icon note
		else
			display dialog "MS Report was updated and needs to add a few packages (internet needed, about a minute)." ¬
				buttons {"Cancel", "Continue"} default button "Continue" with title appTitle with icon note
		end if
		do shell script "/bin/bash " & helper & " install_bg"
		set progress total steps to 5
		set progress completed steps to 0
		set progress description to "Installing MS Report"
		set progress additional description to "Starting"
		set outcome to ""
		repeat 3600 times
			delay 1
			set statusLine to do shell script "/bin/bash " & helper & " install_status"
			if statusLine is "done" then
				set outcome to "done"
				exit repeat
			else if statusLine starts with "failed" then
				set outcome to statusLine
				exit repeat
			else if statusLine contains "|" then
				set AppleScript's text item delimiters to "|"
				set progress completed steps to ((text item 1 of statusLine) as integer)
				set progress additional description to (text item 2 of statusLine)
				set AppleScript's text item delimiters to ""
			end if
		end repeat
		if outcome is not "done" then
			display dialog "The installation did not finish." & return & return & outcome & return & return & ¬
				"Check the internet connection and open MS Report again." buttons {"OK"} default button 1 ¬
				with title appTitle with icon stop
			return
		end if
		set progress completed steps to 5
		set progress additional description to "Opening the window"
	end if

	-- 2. Open the window (it keeps running after this launcher quits) -------------------
	try
		if target is "" then
			do shell script "/bin/bash " & helper & " dashboard"
		else
			do shell script "/bin/bash " & helper & " dashboard " & quoted form of target
		end if
	on error errMsg
		display dialog "MS Report could not open its window." & return & return & errMsg & return & return & ¬
			"If this keeps happening, report it with a screenshot of this message at github.com/Caroldorium/MS-report/issues" buttons {"OK"} default button 1 ¬
			with title appTitle with icon stop
	end try
end startApp
