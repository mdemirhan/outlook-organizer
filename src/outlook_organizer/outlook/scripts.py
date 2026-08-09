from __future__ import annotations

# ruff: noqa: E501

LIST_FOLDERS = r"""
on cleanText(sourceText)
    set valueText to sourceText as text
    set AppleScript's text item delimiters to {return, linefeed, tab}
    set textParts to text items of valueText
    set AppleScript's text item delimiters to " "
    set valueText to textParts as text
    set AppleScript's text item delimiters to ""
    return valueText
end cleanText

on run argv
    set maximumFolderID to (item 1 of argv) as integer
    tell application "Microsoft Outlook"
        set outputText to ""
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
"""


LIST_CALENDARS = r"""
on cleanText(sourceText)
    set valueText to sourceText as text
    set AppleScript's text item delimiters to {return, linefeed, tab}
    set textParts to text items of valueText
    set AppleScript's text item delimiters to " "
    set valueText to textParts as text
    set AppleScript's text item delimiters to ""
    return valueText
end cleanText

on run argv
    set maximumCalendarID to (item 1 of argv) as integer
    tell application "Microsoft Outlook"
        set outputText to ""
        repeat with candidateID from 1 to maximumCalendarID
            try
                set calendarRef to calendar id (candidateID as integer)
                set outputText to outputText & candidateID & tab & my cleanText(name of calendarRef) & tab & (count of calendar events of calendarRef) & linefeed
            end try
        end repeat
        return outputText
    end tell
end run
"""


READ_MESSAGES = r"""
on replaceText(sourceText, searchText, replacementText)
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to searchText
    set textParts to text items of (sourceText as text)
    set AppleScript's text item delimiters to replacementText
    set resultText to textParts as text
    set AppleScript's text item delimiters to previousDelimiters
    return resultText
end replaceText

on cleanField(sourceText)
    set resultText to sourceText as text
    repeat with codePoint in {28, 29, 30, 31}
        set resultText to my replaceText(resultText, character id codePoint, " ")
    end repeat
    return resultText
end cleanField

on padTwo(numberValue)
    set textValue to numberValue as text
    if (length of textValue) < 2 then set textValue to "0" & textValue
    return textValue
end padTwo

on isoDate(dateValue)
    return (year of dateValue as text) & "-" & my padTwo((month of dateValue) as integer) & "-" & my padTwo(day of dateValue) & "T" & my padTwo(hours of dateValue) & ":" & my padTwo(minutes of dateValue) & ":" & my padTwo(seconds of dateValue)
end isoDate

on recipientListText(recipientRefs)
    set groupSeparator to character id 29
    set fieldSeparator to character id 28
    set resultItems to {}
    tell application "Microsoft Outlook"
        repeat with recipientRef in recipientRefs
            set addressRecord to email address of recipientRef
            set addressName to ""
            try
                set addressName to get name of addressRecord
            end try
            set addressValue to ""
            try
                set addressValue to get address of addressRecord
            end try
            if addressName is "" then set addressName to addressValue
            set addressType to "unknown"
            try
                set addressType to (get type of addressRecord) as text
            end try
            set end of resultItems to my cleanField(addressName) & fieldSeparator & my cleanField(addressValue) & fieldSeparator & my cleanField(addressType)
        end repeat
    end tell
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to groupSeparator
    set resultText to resultItems as text
    set AppleScript's text item delimiters to previousDelimiters
    return resultText
end recipientListText

on messageRow(messageRef, folderRef, bodyLimit)
    set unitSeparator to character id 31
    set fieldValues to {}
    tell application "Microsoft Outlook"
        set end of fieldValues to id of messageRef as text
        set exchangeValue to ""
        try
            set exchangeValue to exchange id of messageRef
        end try
        set end of fieldValues to my cleanField(exchangeValue)
        set end of fieldValues to id of folderRef as text
        set end of fieldValues to my cleanField(name of folderRef)
        set end of fieldValues to my cleanField(subject of messageRef)

        set senderRecord to sender of messageRef
        set senderName to ""
        try
            set senderName to get name of senderRecord
        end try
        set senderAddress to ""
        try
            set senderAddress to get address of senderRecord
        end try
        if senderName is "" then set senderName to senderAddress
        set end of fieldValues to my cleanField(senderName)
        set end of fieldValues to my cleanField(senderAddress)
        set end of fieldValues to my recipientListText(every to recipient of messageRef)
        set end of fieldValues to my recipientListText(every cc recipient of messageRef)
        set end of fieldValues to my isoDate(time received of messageRef)
        set end of fieldValues to my cleanField((todo flag of messageRef) as text)

        set categoryNames to {}
        repeat with categoryRef in (get categories of messageRef)
            set end of categoryNames to my cleanField(name of categoryRef)
        end repeat
        set previousDelimiters to AppleScript's text item delimiters
        set AppleScript's text item delimiters to character id 29
        set end of fieldValues to categoryNames as text
        set AppleScript's text item delimiters to previousDelimiters

        set bodyValue to ""
        if bodyLimit > 0 then
            try
                set bodyValue to get plain text content of messageRef
                if (length of bodyValue) > bodyLimit then set bodyValue to text 1 thru bodyLimit of bodyValue
            end try
        end if
        set end of fieldValues to my cleanField(bodyValue)
        set end of fieldValues to ((count of attachments of messageRef) > 0) as text

        set threadGUIDValue to ""
        try
            set threadGUIDValue to «class lOTd» of messageRef as text
        end try
        set end of fieldValues to my cleanField(threadGUIDValue)

        set readValue to true
        try
            set readValue to get is read of messageRef
        end try
        set end of fieldValues to readValue as text

        set repliedValue to false
        try
            set repliedValue to get «class pRpT» of messageRef
        end try
        set end of fieldValues to repliedValue as text
    end tell

    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to unitSeparator
    set resultText to fieldValues as text
    set AppleScript's text item delimiters to previousDelimiters
    return resultText
end messageRow

on run argv
    set folderID to (item 1 of argv) as integer
    set requestedCount to (item 2 of argv) as integer
    set bodyLimit to (item 3 of argv) as integer
    tell application "Microsoft Outlook"
        set folderRef to mail folder id (folderID as integer)
        set messageRefs to every message of folderRef
        set receivedDates to time received of every message of folderRef
        set selectedIndexes to {}

        repeat with rankNumber from 1 to requestedCount
            set bestIndex to 0
            set bestDate to missing value
            repeat with messageIndex from 1 to count of receivedDates
                set numericIndex to messageIndex as integer
                if selectedIndexes does not contain numericIndex then
                    set receivedDate to item numericIndex of receivedDates
                    if receivedDate is not missing value then
                        if bestDate is missing value or receivedDate > bestDate then
                            set bestDate to receivedDate
                            set bestIndex to numericIndex
                        end if
                    end if
                end if
            end repeat
            if bestIndex > 0 then set end of selectedIndexes to bestIndex
        end repeat

        set rows to {}
        repeat with selectedIndex in selectedIndexes
            set numericIndex to selectedIndex as integer
            set end of rows to my messageRow(item numericIndex of messageRefs, folderRef, bodyLimit)
        end repeat
        set previousDelimiters to AppleScript's text item delimiters
        set AppleScript's text item delimiters to character id 30
        set outputText to rows as text
        set AppleScript's text item delimiters to previousDelimiters
        return outputText
    end tell
end run
"""


READ_MESSAGE_BY_ID = (
    READ_MESSAGES.split("on run argv", 1)[0]
    + r"""

on run argv
    set messageID to (item 1 of argv) as integer
    set bodyLimit to (item 2 of argv) as integer
    tell application "Microsoft Outlook"
        set messageRef to message id (messageID as integer)
        set folderRef to folder of messageRef
        return my messageRow(messageRef, folderRef, bodyLimit)
    end tell
end run
"""
)


READ_MESSAGES_IN_FOLDER_ORDER = (
    READ_MESSAGES.split("on run argv", 1)[0]
    + r"""

on run argv
    set folderID to (item 1 of argv) as integer
    set requestedCount to (item 2 of argv) as integer
    set bodyLimit to (item 3 of argv) as integer
    tell application "Microsoft Outlook"
        set folderRef to mail folder id (folderID as integer)
        set availableCount to count of messages of folderRef
        if requestedCount > availableCount then set requestedCount to availableCount
        if requestedCount = 0 then return ""
        set messageRefs to messages 1 thru requestedCount of folderRef
        set rows to {}
        repeat with messageIndex from 1 to requestedCount
            set messageRef to item messageIndex of messageRefs
            set receivedDate to missing value
            try
                set receivedDate to time received of messageRef
            end try
            if receivedDate is not missing value then
                set end of rows to my messageRow(messageRef, folderRef, bodyLimit)
            end if
        end repeat
        set previousDelimiters to AppleScript's text item delimiters
        set AppleScript's text item delimiters to character id 30
        set outputText to rows as text
        set AppleScript's text item delimiters to previousDelimiters
        return outputText
    end tell
end run
"""
)


READ_MESSAGES_IN_FOLDER_WINDOW = (
    READ_MESSAGES.split("on run argv", 1)[0]
    + r"""

on run argv
    set folderID to (item 1 of argv) as integer
    set startOffsetSeconds to (item 2 of argv) as integer
    set endOffsetSeconds to (item 3 of argv) as integer
    set desiredReadState to item 4 of argv as text
    set requestedCount to (item 5 of argv) as integer
    set bodyLimit to (item 6 of argv) as integer
    set nowValue to current date
    set windowStart to nowValue + startOffsetSeconds
    set windowEnd to nowValue + endOffsetSeconds

    tell application "Microsoft Outlook"
        set folderRef to mail folder id (folderID as integer)
        set messageRefs to every message of folderRef whose time received >= windowStart and time received < windowEnd
        set rows to {}
        repeat with messageRef in messageRefs
            set includeMessage to true
            if desiredReadState is not "all" then
                set messageIsRead to true
                try
                    set messageIsRead to get is read of messageRef
                end try
                if desiredReadState is "unread" and messageIsRead then
                    set includeMessage to false
                else if desiredReadState is "read" and not messageIsRead then
                    set includeMessage to false
                end if
            end if
            if includeMessage then
                set end of rows to my messageRow(messageRef, folderRef, bodyLimit)
                if (count of rows) >= requestedCount then exit repeat
            end if
        end repeat
        set previousDelimiters to AppleScript's text item delimiters
        set AppleScript's text item delimiters to character id 30
        set outputText to rows as text
        set AppleScript's text item delimiters to previousDelimiters
        return outputText
    end tell
end run
"""
)


LIST_MESSAGE_DATES = r"""
on padTwo(numberValue)
    set textValue to numberValue as text
    if (length of textValue) < 2 then set textValue to "0" & textValue
    return textValue
end padTwo

on isoDate(dateValue)
    return (year of dateValue as text) & "-" & my padTwo((month of dateValue) as integer) & "-" & my padTwo(day of dateValue) & "T" & my padTwo(hours of dateValue) & ":" & my padTwo(minutes of dateValue) & ":" & my padTwo(seconds of dateValue)
end isoDate

on run argv
    set folderID to (item 1 of argv) as integer
    set requestedCount to (item 2 of argv) as integer
    tell application "Microsoft Outlook"
        set folderRef to mail folder id (folderID as integer)
        set windowDays to 7
        set messageRefs to {}
        repeat
            set cutoffDate to (current date) - (windowDays * days)
            set messageRefs to every message of folderRef whose time received ≥ cutoffDate
            if (count of messageRefs) ≥ requestedCount then exit repeat
            if windowDays ≥ 3650 then exit repeat
            if windowDays < 30 then
                set windowDays to 30
            else if windowDays < 90 then
                set windowDays to 90
            else
                set windowDays to windowDays * 2
            end if
        end repeat
        set rows to {}
        repeat with messageIndex from 1 to count of messageRefs
            set messageRef to item messageIndex of messageRefs
            try
                set receivedDate to time received of messageRef
                if receivedDate is not missing value then
                    set end of rows to (id of messageRef as text) & (character id 31) & my isoDate(receivedDate)
                end if
            end try
        end repeat
        set previousDelimiters to AppleScript's text item delimiters
        set AppleScript's text item delimiters to character id 30
        set outputText to rows as text
        set AppleScript's text item delimiters to previousDelimiters
        return outputText
    end tell
end run
"""


READ_MESSAGES_BY_IDS = (
    READ_MESSAGES.split("on run argv", 1)[0]
    + r"""

on run argv
    set bodyLimit to (item 1 of argv) as integer
    tell application "Microsoft Outlook"
        set rows to {}
        repeat with argumentIndex from 2 to count of argv
            set messageID to (item argumentIndex of argv) as integer
            try
                set messageRef to message id (messageID as integer)
                set messageFolderRef to folder of messageRef
                set end of rows to my messageRow(messageRef, messageFolderRef, bodyLimit)
            end try
        end repeat
        set previousDelimiters to AppleScript's text item delimiters
        set AppleScript's text item delimiters to character id 30
        set outputText to rows as text
        set AppleScript's text item delimiters to previousDelimiters
        return outputText
    end tell
end run
"""
)


READ_CALENDAR_EVENTS = r"""
on replaceText(sourceText, searchText, replacementText)
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to searchText
    set textParts to text items of (sourceText as text)
    set AppleScript's text item delimiters to replacementText
    set resultText to textParts as text
    set AppleScript's text item delimiters to previousDelimiters
    return resultText
end replaceText

on cleanField(sourceText)
    set resultText to sourceText as text
    repeat with codePoint in {28, 29, 30, 31}
        set resultText to my replaceText(resultText, character id codePoint, " ")
    end repeat
    return resultText
end cleanField

on padTwo(numberValue)
    set textValue to numberValue as text
    if (length of textValue) < 2 then set textValue to "0" & textValue
    return textValue
end padTwo

on isoDate(dateValue)
    return (year of dateValue as text) & "-" & my padTwo((month of dateValue) as integer) & "-" & my padTwo(day of dateValue) & "T" & my padTwo(hours of dateValue) & ":" & my padTwo(minutes of dateValue) & ":" & my padTwo(seconds of dateValue)
end isoDate

on attendeeText(attendeeRef)
    set fieldSeparator to character id 28
    tell application "Microsoft Outlook"
        set addressRecord to email address of attendeeRef
        set addressName to ""
        try
            set addressName to get name of addressRecord
        end try
        set addressValue to ""
        try
            set addressValue to get address of addressRecord
        end try
        if addressName is "" then set addressName to addressValue
        return my cleanField(addressName) & fieldSeparator & my cleanField(addressValue) & fieldSeparator & my cleanField((type of attendeeRef) as text) & fieldSeparator & my cleanField((status of attendeeRef) as text)
    end tell
end attendeeText

on run argv
    set calendarID to (item 1 of argv) as integer
    set daysBehind to (item 2 of argv) as integer
    set daysAhead to (item 3 of argv) as integer
    set bodyLimit to (item 4 of argv) as integer
    tell application "Microsoft Outlook"
        set calendarRef to calendar id (calendarID as integer)
        set windowStart to (current date) - (daysBehind * days)
        set windowEnd to (current date) + (daysAhead * days)
        set eventRefs to every calendar event of calendarRef whose start time ≥ windowStart and start time ≤ windowEnd
        set rows to {}
        repeat with eventRef in eventRefs
            set fields to {}
            set end of fields to id of eventRef as text
            set exchangeValue to ""
            try
                set exchangeValue to exchange id of eventRef
            end try
            set end of fields to my cleanField(exchangeValue)
            set end of fields to calendarID as text
            set end of fields to my cleanField(subject of eventRef)
            set end of fields to my isoDate(start time of eventRef)
            set end of fields to my isoDate(end time of eventRef)
            set end of fields to my cleanField(location of eventRef)
            set end of fields to my cleanField(organizer of eventRef)
            set end of fields to (all day flag of eventRef) as text
            set end of fields to my cleanField((free busy status of eventRef) as text)
            set end of fields to (is private of eventRef) as text

            set categoryNames to {}
            repeat with categoryRef in (get categories of eventRef)
                set end of categoryNames to my cleanField(name of categoryRef)
            end repeat
            set previousDelimiters to AppleScript's text item delimiters
            set AppleScript's text item delimiters to character id 29
            set end of fields to categoryNames as text

            set attendeeItems to {}
            repeat with attendeeRef in every attendee of eventRef
                set end of attendeeItems to my attendeeText(attendeeRef)
            end repeat
            set AppleScript's text item delimiters to character id 29
            set end of fields to attendeeItems as text
            set AppleScript's text item delimiters to previousDelimiters

            set bodyValue to ""
            if bodyLimit > 0 then
                try
                    set bodyValue to get plain text content of eventRef
                    if (length of bodyValue) > bodyLimit then set bodyValue to text 1 thru bodyLimit of bodyValue
                end try
            end if
            set end of fields to my cleanField(bodyValue)

            set AppleScript's text item delimiters to character id 31
            set end of rows to fields as text
            set AppleScript's text item delimiters to previousDelimiters
        end repeat
        set AppleScript's text item delimiters to character id 30
        set outputText to rows as text
        set AppleScript's text item delimiters to ""
        return outputText
    end tell
end run
"""


ENSURE_MAIL_FOLDER = r"""
on cleanText(sourceText)
    set valueText to sourceText as text
    set AppleScript's text item delimiters to {return, linefeed, tab}
    set textParts to text items of valueText
    set AppleScript's text item delimiters to " "
    set valueText to textParts as text
    set AppleScript's text item delimiters to ""
    return valueText
end cleanText

on run argv
    set inboxID to (item 1 of argv) as integer
    set desiredName to item 2 of argv
    tell application "Microsoft Outlook"
        set inboxRef to mail folder id (inboxID as integer)
        set accountRef to account of inboxRef
        set existingFolders to every mail folder of accountRef whose name is desiredName
        if (count of existingFolders) > 0 then
            set folderRef to item 1 of existingFolders
            return (id of folderRef as text) & tab & my cleanText(name of folderRef) & tab & "existing"
        end if
        set folderRef to make new mail folder at accountRef with properties {name:desiredName}
        return (id of folderRef as text) & tab & my cleanText(name of folderRef) & tab & "created"
    end tell
end run
"""


APPLY_MAIL_STATE = r"""
on run argv
    set messageID to (item 1 of argv) as integer
    set categoryText to item 2 of argv
    set desiredFlag to item 3 of argv
    set targetFolderID to (item 4 of argv) as integer

    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to character id 29
    set categoryNames to text items of categoryText
    set AppleScript's text item delimiters to previousDelimiters

    tell application "Microsoft Outlook"
        set messageRef to message id (messageID as integer)
        set categoryRefs to {}
        repeat with categoryNameRef in categoryNames
            set categoryName to categoryNameRef as text
            if categoryName is not "" then
                try
                    set categoryRef to first category whose name is categoryName
                on error
                    set categoryRef to make new category with properties {name:categoryName}
                end try
                set end of categoryRefs to categoryRef
            end if
        end repeat
        set categories of messageRef to categoryRefs

        if desiredFlag is "flagged" then
            set todo flag of messageRef to not completed
        else if desiredFlag is "completed" then
            set todo flag of messageRef to completed
        else
            set todo flag of messageRef to not flagged
        end if

        if targetFolderID > 0 then
            set targetFolder to mail folder id (targetFolderID as integer)
            move messageRef to targetFolder
        end if
        return "ok"
    end tell
end run
"""


APPLY_MAIL_STATES = r"""
on replaceText(sourceText, searchText, replacementText)
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to searchText
    set textParts to text items of (sourceText as text)
    set AppleScript's text item delimiters to replacementText
    set resultText to textParts as text
    set AppleScript's text item delimiters to previousDelimiters
    return resultText
end replaceText

on cleanField(sourceText)
    set resultText to sourceText as text
    repeat with codePoint in {30, 31}
        set resultText to my replaceText(resultText, character id codePoint, " ")
    end repeat
    return resultText
end cleanField

on applyOne(messageID, categoryText, desiredFlag, targetFolderID)
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to character id 29
    set categoryNames to text items of categoryText
    set AppleScript's text item delimiters to previousDelimiters

    tell application "Microsoft Outlook"
        set messageRef to message id (messageID as integer)
        set categoryRefs to {}
        repeat with categoryNameRef in categoryNames
            set categoryName to categoryNameRef as text
            if categoryName is not "" then
                try
                    set categoryRef to first category whose name is categoryName
                on error
                    set categoryRef to make new category with properties {name:categoryName}
                end try
                set end of categoryRefs to categoryRef
            end if
        end repeat
        set categories of messageRef to categoryRefs

        if desiredFlag is "flagged" then
            set todo flag of messageRef to not completed
        else if desiredFlag is "completed" then
            set todo flag of messageRef to completed
        else
            set todo flag of messageRef to not flagged
        end if

        if targetFolderID > 0 then
            set targetFolder to mail folder id (targetFolderID as integer)
            move messageRef to targetFolder
        end if
    end tell
end applyOne

on run argv
    set rows to {}
    set argumentIndex to 1
    repeat while argumentIndex ≤ (count of argv)
        set messageIDText to item argumentIndex of argv
        set categoryText to item (argumentIndex + 1) of argv
        set desiredFlag to item (argumentIndex + 2) of argv
        set targetFolderID to (item (argumentIndex + 3) of argv) as integer
        try
            my applyOne((messageIDText as integer), categoryText, desiredFlag, targetFolderID)
            set end of rows to messageIDText & (character id 31) & "applied" & (character id 31) & ""
        on error errorMessage number errorNumber
            set end of rows to messageIDText & (character id 31) & "failed" & (character id 31) & my cleanField((errorNumber as text) & ": " & errorMessage)
        end try
        set argumentIndex to argumentIndex + 4
    end repeat
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to character id 30
    set outputText to rows as text
    set AppleScript's text item delimiters to previousDelimiters
    return outputText
end run
"""
