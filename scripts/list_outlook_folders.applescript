on cleanText(sourceText)
	set previousDelimiters to AppleScript's text item delimiters
	set AppleScript's text item delimiters to {return, linefeed, tab}
	set textParts to text items of (sourceText as text)
	set AppleScript's text item delimiters to " "
	set resultText to textParts as text
	set AppleScript's text item delimiters to previousDelimiters
	return resultText
end cleanText

on run argv
	set maximumFolderID to 1000
	if (count of argv) > 0 then set maximumFolderID to (item 1 of argv) as integer

	tell application "Microsoft Outlook"
		set outputText to "ID" & tab & "Folder" & tab & "Parent" & tab & "Messages" & linefeed

		repeat with candidateID from 1 to maximumFolderID
			try
				set folderRef to mail folder id (candidateID as integer)
				set folderName to my cleanText(name of folderRef)
				set parentName to ""

				try
					set parentRef to container of folderRef
					if parentRef is not missing value then set parentName to my cleanText(name of parentRef)
				end try

				set outputText to outputText & candidateID & tab & folderName & tab & parentName & tab & (count of messages of folderRef) & linefeed
			end try
		end repeat

		return outputText
	end tell
end run
