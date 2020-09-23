<div align="center">

## Simple Email Send


</div>

### Description

While browsing through the files on this site, I noticed that most listings that dealt with sending e-mail using VB used the Winsock Control. I recently wrote a small demo for a customer that reads a database file, and then sends it to an email address. This code requires that you use the MAPISession Control as well the MAPIMessage control. You must also have a mail service installed on your system (Exchange,Outlook, ect.). Other than those requirements, the code is very easy to understand.
 
### More Info
 
Mail Users Name, email address, and any file attachments you wish to include.

Make sure that you have the MAPISession control and the MAPIMessages control on the tool box and on the form.

' 

----

Edited 11/08/1999 by Matmonk 

----

I have received a few questions asking how to send the email to multiple addresses at one time. The simplest way that I know of is when you assign a value to the "TO:" string, simply separate the multiple address with a semicolon ";". This is the microsoft standard way that works in Outlook and seems to work in other email system. An example: strToString = "smtp:JJones@somewhere.com;smtp:SSmith@nowhere.com;smtp:Whoever@microsoft.com". I hope this helps those of you who have raised this question. I do not yet have an answer on how to send multiple file attachments, but if I come across one, I will post it here. Cheers.

----

Edited 11/08/1999 by Matmonk 

----



None that I have found yet.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matmonk](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matmonk.md)
**Level**          |Unknown
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matmonk-simple-email-send__1-1862/archive/master.zip)





### Source Code

```
Private Sub cmdSendSummary_Click()
' this command button is used to start a MAPI session, log on the the
' mail service, attach the created check summary text file to a new
' message, send the message and then close the session
' declare local variables here
 Dim strUserId As String
 Dim strPassword As String
 Dim strFileName As String
 Dim strFilePath As String
' set the mouse pointer to indicate the app is busy
 Screen.MousePointer = vbHourglass
' set the values for the file name and the file path
 strFileName = "" ' this is where you would put any file attachments
 strFilePath = App.Path & "\"
' set the user name and password properties on the session control
 mapiLogOn.UserName = "JJones" ' network user name and password !
 mapiLogOn.Password = "******"
' start a new email session
 mapiLogOn.SignOn
 Do While mapiLogOn.SessionID = 0
  DoEvents ' need to wait until the new session is created
 Loop
'create a new message and address it
 MAPIMessages1.SessionID = mapiLogOn.SessionID
 MAPIMessages1.Compose
 MAPIMessages1.RecipDisplayName = "Jones,John"
 MAPIMessages1.AddressResolveUI = True
 MAPIMessages1.ResolveName
 MAPIMessages1.RecipAddress = "smtp:someone@somewhere.com"
' note that I prefixed the address with "smtp". This is required by exchange
' server, or it does not know what service to use for outgoing mail.
 MAPIMessages1.MsgSubject = "Test of the Email function"
 MAPIMessages1.MsgNoteText = " This is a test of the email function, if you" _
  & "receive this then the program has worked successfully." & vbCrLf
' attaching the file
 MAPIMessages1.AttachmentPosition = Len(MAPIMessages1.MsgNoteText) - 1
' the line above places the attachment at the end of the text.
 MAPIMessages1.AttachmentPathName = strFilePath & strFileName
' now send the message
 MAPIMessages1.Send False
 mapiLogOn.SignOff
 MsgBox "File sent to specified receiptent."
' now set the mouse pointer back to normal
 Screen.MousePointer = vbNormal
End Sub
```

