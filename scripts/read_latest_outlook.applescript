on replaceText(sourceText, searchText, replacementText)
	set previousDelimiters to AppleScript's text item delimiters
	set AppleScript's text item delimiters to searchText
	set textParts to text items of (sourceText as text)
	set AppleScript's text item delimiters to replacementText
	set resultText to textParts as text
	set AppleScript's text item delimiters to previousDelimiters
	return resultText
end replaceText

on cleanText(sourceText)
	set resultText to sourceText as text
	set resultText to my replaceText(resultText, return, " ")
	set resultText to my replaceText(resultText, linefeed, " ")
	set resultText to my replaceText(resultText, tab, " ")
	return resultText
end cleanText

on run argv
	set requestedCount to 10
	set maximumFolderID to 1000
	if (count of argv) > 0 then set requestedCount to (item 1 of argv) as integer
	if (count of argv) > 1 then set maximumFolderID to (item 2 of argv) as integer

	tell application "Microsoft Outlook"
		-- Migrated Exchange profiles can appear in Outlook's navigation pane while
		-- their account and folder properties return "missing value". Folder lookup
		-- by Outlook's own integer ID still returns fully scriptable objects.
		set targetFolder to missing value
		set targetMessageCount to -1

		repeat with candidateID from 1 to maximumFolderID
			try
				set candidateFolder to mail folder id (candidateID as integer)
				set candidateName to name of candidateFolder
				if candidateName is "Inbox" or candidateName is "Gelen Kutusu" then
					set candidateMessageCount to count of messages of candidateFolder
					if candidateMessageCount > targetMessageCount then
						set targetFolder to candidateFolder
						set targetMessageCount to candidateMessageCount
					end if
				end if
			end try
		end repeat

		if targetFolder is missing value then error "No scriptable Inbox or Gelen Kutusu folder was found."

		set messageRefs to every message of targetFolder
		set receivedDates to time received of every message of targetFolder
		set selectedIndexes to {}

		repeat with rankNumber from 1 to requestedCount
			set bestIndex to 0
			set bestDate to missing value

			repeat with messageIndex from 1 to count of receivedDates
				if selectedIndexes does not contain messageIndex then
					set receivedDate to item messageIndex of receivedDates
					if receivedDate is not missing value then
						if bestDate is missing value or receivedDate > bestDate then
							set bestDate to receivedDate
							set bestIndex to messageIndex
						end if
					end if
				end if
			end repeat

			if bestIndex > 0 then set end of selectedIndexes to bestIndex
		end repeat

		set outputText to "Folder" & tab & (name of targetFolder) & tab & targetMessageCount & linefeed

		repeat with selectedIndex in selectedIndexes
			set messageIndex to selectedIndex as integer
			set messageRef to item messageIndex of messageRefs
			set subjectText to my cleanText(subject of messageRef)

			set senderRecord to sender of messageRef
			set senderText to my cleanText(name of senderRecord) & " <" & my cleanText(address of senderRecord) & ">"

			set recipientTexts to {}
			set recipientRefs to every to recipient of messageRef
			repeat with recipientRef in recipientRefs
				set addressRecord to email address of recipientRef
				set recipientName to my cleanText(name of addressRecord)
				set recipientAddress to my cleanText(address of addressRecord)
				if recipientName is "" then
					set end of recipientTexts to recipientAddress
				else
					set end of recipientTexts to recipientName & " <" & recipientAddress & ">"
				end if
			end repeat

			set previousDelimiters to AppleScript's text item delimiters
			set AppleScript's text item delimiters to "; "
			set toText to recipientTexts as text
			set AppleScript's text item delimiters to previousDelimiters

			set outputText to outputText & subjectText & tab & senderText & tab & toText & linefeed
		end repeat

		return outputText
	end tell
end run
