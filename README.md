# Intro
MS Outlook COM : Find and Extract attachments from Outlook folder
https://www.autohotkey.com/boards/viewtopic.php?f=6&t=71377
I wrote this code to find and extract attachments to specific folder

# Requirment
1) . MS Outlook installed
2) . MS Outlook running, do not ask why

# How it work
1) . It look for specific subject and finds email
2) . It look for specific attachments name in found email
3) . It will now look into subfolder

# How to make it fast
1) . Make rule to save email specific email to different folder
2) . Email older then three months get deleted, coz the less the folder has will be faster to lookup
or
3) . make inverse loop so it will look for recent to past order and function `return` when first attachment found

```autohotkey
	Loop % folder.items.count
		thisattachment := attachments.Item(folder.items.count - a_index)
```

# Example
```autohotkey
ExtractOutlookAttachments("Your@emailaddress.com\inbox\Subfolder","Some Subject","AttachmentName", a_desktop)
```
