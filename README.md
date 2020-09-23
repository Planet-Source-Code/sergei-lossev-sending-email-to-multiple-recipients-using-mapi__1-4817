<div align="center">

## Sending Email to multiple recipients using MAPI


</div>

### Description

It seems that a lot of people are having problems sending Email from VB to multiple people. Here's a procedure that should fix that.
 
### More Info
 
EmailAddress is of the form "first@first.com;second@second.com" and etc.

Subject and MessageText are self explainatory.

You are gonna need a form with the two MAPI controls, named like in the code.

The rest you should be able to get by yourself.

Should work. It does on my computer :o)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sergei Lossev](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sergei-lossev.md)
**Level**          |Unknown
**User Rating**    |4.8 (43 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sergei-lossev-sending-email-to-multiple-recipients-using-mapi__1-4817/archive/master.zip)





### Source Code

```
Public Sub SendEmail(sEmailAddress As String,sSubject as string, sMessageText as string)
  Dim sEmailExtracted As String
  Dim sEmailLeft As String
  Dim iRecipCount As Integer
  If Trim(sEmailAddress) = "" Then
      Goto SendMail_End
  End If
  sEmailLeft = Trim(sEmailAddress)
  ' set the mouse pointer to indicate the app is busy
  Screen.MousePointer = vbHourglass
  MAPIlogon.SignOn
  Do While MAPIlogon.SessionID = 0
    DoEvents ' need to wait until the new session is created
  Loop
    With MAPIMessages1
      .MsgIndex = -1
      .SessionID = MAPIlogon.SessionID
      While sEmailLeft <> ""
        If InStr(1, sEmailLeft, ";") = 0 Then
          sEmailExtracted = sEmailLeft
          sEmailLeft = ""
        Else
          sEmailExtracted = Left(sEmailLeft, InStr(1, sEmailLeft, ";") - 1)
          sEmailLeft = Right(sEmailLeft, Len(sEmailLeft) - InStr(1, sEmailLeft, ";"))
        End If
        .RecipIndex = iRecipCount
        If iRecipCount = 0 Then
          .RecipType = mapToList
        Else
          .RecipType = mapCcList
        End If
        .RecipAddress = sEmailExtracted
        .ResolveName
        iRecipCount = iRecipCount + 1
      Wend
      If iRecipCount = 0 Then GoTo SendMail_End
      .MsgSubject = sSubject
      .MsgNoteText = sMessageText
      .Send
    End With
    MAPIlogon.SignOff
SendMail_End:
  Screen.MousePointer = vbNormal
  Exit Sub
End Sub
```

