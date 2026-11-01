-- SMR Filler. Double-click, pick the workbook, pick the txt, done dialog.
-- The bundled smr_filler.py lives at Contents/Resources/smr_filler.py.
-- On open, the app tries to download the release copy. The window title
-- is the VERSION line of the file that actually runs.

on run
	try
		set bundledScript to POSIX path of (path to resource "smr_filler.py")
	on error
		display dialog "This app is missing its bundled smr_filler.py. Re-download the SMR Filler package and try again." buttons {"OK"} default button "OK" with icon stop with title "SMR Filler"
		return
	end try

	set homePosix to POSIX path of (path to home folder)
	if homePosix ends with "/" then
		set supportDir to homePosix & "Library/Application Support/smr-filler"
	else
		set supportDir to homePosix & "/Library/Application Support/smr-filler"
	end if
	set supportFile to supportDir & "/smr_filler.py"
	set tmpFile to "/tmp/smr_filler_update.py"
	set updateURL to "https://raw.githubusercontent.com/jwatson7399/smr/release/filler/smr_filler.py"

	try
		do shell script "curl -fsSL --max-time 10 " & quoted form of updateURL & " -o " & quoted form of tmpFile & " && grep -q 'VERSION = ' " & quoted form of tmpFile & " && mkdir -p " & quoted form of supportDir & " && mv " & quoted form of tmpFile & " " & quoted form of supportFile & " || rm -f " & quoted form of tmpFile
	end try

	set fillerScript to bundledScript
	try
		do shell script "test -f " & quoted form of supportFile
		set fillerScript to supportFile
	end try

	set appVersion to "unknown"
	try
		set verLine to do shell script "grep 'VERSION = ' " & quoted form of fillerScript & " | head -1"
		set AppleScript's text item delimiters to "\""
		set appVersion to text item 2 of verLine
		set AppleScript's text item delimiters to ""
	end try
	if appVersion is "" then set appVersion to "unknown"

	try
		do shell script "python3 -c 'import openpyxl'"
	on error
		display dialog "The filler isn't set up yet on this Mac." & return & return & "Open the SMR Filler folder and double-click '1. Setup (run once).command' first." buttons {"OK"} default button "OK" with icon stop with title "SMR Filler " & appVersion
		return
	end try

	try
		set wbFile to choose file with prompt "Select the SMR billing workbook (.xlsm):"
	on error number -128
		return
	end try
	set wbPath to POSIX path of wbFile

	try
		set txtFile to choose file with prompt "Select the data file (.txt) from the phone app:"
	on error number -128
		return
	end try
	set txtPath to POSIX path of txtFile

	try
		set runCmd to "python3 " & quoted form of fillerScript & " " & quoted form of wbPath & " " & quoted form of txtPath & " 2>&1"
		set fillerOutput to do shell script runCmd
	on error errMsg
		display dialog "Something went wrong while filling the workbook:" & return & return & errMsg buttons {"OK"} default button "OK" with icon stop with title "SMR Filler " & appVersion
		return
	end try

	set outPath to do shell script "dir=$(dirname " & quoted form of wbPath & "); base=$(basename " & quoted form of wbPath & "); name=\"${base%.*}\"; ext=\"${base##*.}\"; echo \"$dir/${name}_filled.$ext\""

	-- The filler prints "CHECK: " lines last: things to look at before submitting.
	set checkLines to {}
	repeat with p in paragraphs of fillerOutput
		set lineText to p as text
		if lineText starts with "CHECK: " and (length of lineText) > 7 then set end of checkLines to text 8 thru -1 of lineText
	end repeat

	do shell script "open -R " & quoted form of outPath
	set doneMsg to "Done!" & return & return & "Filled workbook saved as:" & return & outPath
	if (count of checkLines) is 0 then
		display dialog doneMsg buttons {"OK"} default button "OK" with title "SMR Filler " & appVersion
	else
		set doneMsg to doneMsg & return & return & "Please check before submitting:"
		repeat with i from 1 to count of checkLines
			if i > 8 then
				set doneMsg to doneMsg & return & "...and " & ((count of checkLines) - 8) & " more."
				exit repeat
			end if
			set doneMsg to doneMsg & return & return & "- " & item i of checkLines
		end repeat
		display dialog doneMsg buttons {"OK"} default button "OK" with icon caution with title "SMR Filler " & appVersion
	end if
end run
