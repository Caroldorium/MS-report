-- MS Report.app: installs the tools once, then opens the MS Report window (dashboard.py).
-- All the work is done by msreport.sh inside the app; this file only shows the install dialogs.

property appTitle : "MS Report"

on run
	set helper to quoted form of (POSIX path of (path to resource "msreport.sh"))

	-- 1. One-time install (and again after an update that needs new packages) ----------
	if (do shell script "/bin/bash " & helper & " ready") is not "ready" then
		set firstTime to (do shell script "/bin/bash " & helper & " installed_before") is "no"
		if firstTime then
			display dialog "Welcome to MS Report." & return & return & ¬
				"Before the first use, MS Report downloads the tools it needs (about 600 MB, internet needed). " & ¬
				"This takes 2-5 minutes and happens only once." & return & return & ¬
				"You will see a message when it is finished." buttons {"Cancel", "Install"} ¬
				default button "Install" with title appTitle with icon note
		else
			display dialog "MS Report was updated and needs to add a few packages (internet needed, about a minute)." ¬
				buttons {"Cancel", "Continue"} default button "Continue" with title appTitle with icon note
		end if
		display notification "Installing - this takes a few minutes." with title appTitle
		try
			with timeout of 3600 seconds
				do shell script "/bin/bash " & helper & " install"
			end timeout
		on error errMsg
			display dialog "The installation did not finish." & return & return & errMsg & return & return & ¬
				"Check the internet connection and open MS Report again." buttons {"OK"} default button 1 ¬
				with title appTitle with icon stop
			return
		end try
	end if

	-- 2. Open the window (it keeps running after this launcher quits) -------------------
	try
		do shell script "/bin/bash " & helper & " dashboard"
	on error errMsg
		display dialog "MS Report could not open its window." & return & return & errMsg & return & return & ¬
			"If this keeps happening, send a screenshot of this message to Carlos." buttons {"OK"} default button 1 ¬
			with title appTitle with icon stop
	end try
end run
