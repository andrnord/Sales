Attribute VB_Name = "ReplyLastEmail"
Sub ReplyToLastEmail()
    Dim oMail As MailItem
    Dim oReply As MailItem
    Dim Sender As String, emailAdress As String, eFirst As String, eLast As String
    Dim userName As String
    Dim firstname As String, lastname As String, tempname As String
    Dim Greeting As String
    Dim nameArray As Variant, mailArray As Variant
    
    Dim Recipient As Object
     
    Select Case Application.ActiveWindow.Class
           Case olInspector
                Set oMail = ActiveInspector.CurrentItem
           Case olExplorer
                Set oMail = ActiveExplorer.Selection.item(1)
    End Select
    
    Greeting = "Hi "
    Sender = oMail.Sender
    RecipientCount = oMail.Recipients.Count
    
    Set Recipient = CreateObject("System.Collections.ArrayList")
    
    For i = 1 To RecipientCount
        Recipient.Add oMail.Recipients(i)
        
        If InStr(oMail.Recipients(i).Name, "@") = 0 And InStr(oMail.Sender.Address, "Exchange") <> 0 Then
            GoTo Jump
        End If
    Next

    userName = Application.GetNamespace("MAPI").CurrentUser
    
    If Sender = userName Then
        Sender = Recipient(0)
        emailAdress = oMail.Recipients.item(1).Address
        nameArray = Split(Split(emailAdress, "@")(0), ".")
        sString = InStr(1, Sender, "@")
        
    Else
        emailAdress = oMail.Sender.Address
        nameArray = Split(Split(emailAdress, "@")(0), ".")
        sString = InStr(1, Sender, "@")
        
    End If
        
    If oMail.SenderEmailType = "SMTP" And sString = 0 Then
        On Error Resume Next
        emailAdress = Split(emailAdress, "@")(0)
        eFirst = nameArray(0)
        If nameArray > 0 Then eLast = nameArray(1)
        firstname = Split(Sender, " ")(0)
        lastname = Split(Sender, " ")(1)
        
        If firstname <> StrConv(eFirst, vbProperCase) And UBound(nameArray) > 0 Then
            tempname = firstname
            firstname = lastname
            lastname = tempname
            Greeting = "Dear "
        End If
        
    ElseIf sString <> 0 Then
        emailAdress = Split(emailAdress, "@")(0)
        eFirst = Split(emailAdress, ".")(0)
        eLast = Split(emailAdress, ".")(1)
        firstname = StrConv(eFirst, vbProperCase)
        lastname = StrConv(eLast, vbProperCase)
        
    Else
        nameArray = Split(Sender, " ")
        'Greeting = "Hei "
        
        If UBound(nameArray) > 2 Then
        firstname = nameArray(0) & " " & nameArray(1)
        lastname = nameArray(UBound(nameArray))
        
        Else
            firstname = nameArray(0)
            lastname = nameArray(UBound(nameArray))
        End If
        
    End If
                 
   
    Set oReply = oMail.ReplyAll
    
    Dim currentFont As String
    currentFont = oMail.HTMLBody
 
    With oReply
        '.CC = ""
        If InStr(currentFont, "font-size:11pt") > 0 Then
            .HTMLBody = "<p style=""font-family:Calibri (Body)"">" & Greeting & firstname & "," & r1 & strTo & "</p>" & .HTMLBody 'strbody & "<br>" & .HTMLBody
        Else
            .HTMLBody = "<p style=""font-family:Calibri (Body); font-size:11.5pt;"">" & Greeting & firstname & "," & r1 & strTo & "</p>" & .HTMLBody 'strbody & "<br>" & .HTMLBody
        .Display
        End If
    End With
    
Exit Sub
    
Jump:

Dim oForward As MailItem
Set oForward = oMail.Forward
Dim forwarEmailAddress As String

Dim mailBody As String
mailBody = oMail.Body

Dim sSubject As String
sSubject = oMail.Subject

Dim searchTo As String
Dim seachFrom As String
Dim seachCopy As String

If InStr(1, mailBody, "Fra: ") <> 0 Then
    'MsgBox mailBody
    mailBody = Split(Split(mailBody, Right("Emne: ", Len("Emne: ") - 3))(0), "Fra: ")(1)
    searchTo = "Til: "
    seachFrom = "Fra: "
    searchCopy = "Kopi: "
Else
    mailBody = Split(Split(mailBody, Right("Subject: ", Len("Subject: ") - 3))(0), "From: ")(1)
    searchTo = "To: "
    seachFrom = "From: "
    searchCopy = "Cc: "
End If

forwarEmailAddress = Split(Split(mailBody, "<")(1), ">")(0)

Recipient.Clear
If InStr(1, mailBody, searchCopy) = 0 Then
    mailBody = Split(mailBody, searchTo)(1)
Else
    mailBody = Split(mailBody, searchCopy)(1)
End If

For i = 0 To Len(mailBody) - Len(Replace(mailBody, "<", "")) - 1
    Recipient.Add Mid(Split(Split(mailBody, "<")(i + 1), ">")(0), 1, Len(mailBody))
Next
    
    userName = Application.GetNamespace("MAPI").CurrentUser
    On Error Resume Next
      
    If Sender = userName Then
        Sender = forwarEmailAddress
        nameArray = Split(Split(emailAdress, "@")(0), ".")
        sString = InStr(1, Sender, "@")
        
    Else
        emailAdress = forwarEmailAddress
        nameArray = Split(Split(emailAdress, "@")(0), ".")
        sString = InStr(1, emailAdress, "@")
        
    End If
    
    If oMail.SenderEmailType = "SMTP" And sString = 0 Then
        On Error Resume Next
        emailAdress = Split(emailAdress, "@")(0)
        eFirst = nameArray(0)
        If nameArray > 0 Then eLast = nameArray(1)
        firstname = Split(Sender, " ")(0)
        lastname = Split(Sender, " ")(1)
        
        If firstname <> StrConv(eFirst, vbProperCase) And UBound(nameArray) > 0 Then
            tempname = firstname
            firstname = lastname
            lastname = tempname
            Greeting = "Dear "
        End If
        
    ElseIf sString <> 0 Then
        emailAdress = Split(forwarEmailAddress, "@")(0)
        eFirst = Split(emailAdress, ".")(0)
        eLast = Split(emailAdress, ".")(1)
        firstname = StrConv(eFirst, vbProperCase)
        lastname = StrConv(eLast, vbProperCase)
        
    Else
        nameArray = Split(forwarEmailAddress, " ")
        Greeting = "Hi "
        
        If UBound(nameArray) > 2 Then
        firstname = nameArray(0) & " " & nameArray(1)
        lastname = nameArray(UBound(nameArray))
        
        Else
            firstname = nameArray(0)
            lastname = nameArray(UBound(nameArray))
        End If
        
    End If
    
Dim ccList As String
Dim item As Variant

For Each item In Recipient
    If InStr(1, ccList, Split(item, "mailto:")(1)) = 0 Then
        ccList = ccList & ";" & item
    End If
Next item

With oForward
    .To = forwarEmailAddress
    .CC = ccList

    If InStr(currentFont, "font-size:11pt") > 0 Then
        .HTMLBody = "<p style=""font-family:Calibri (Body)"">" & Greeting & firstname & "," & r1 & strTo & "</p>" & .HTMLBody 'strbody & "<br>" & .HTMLBody
    Else
        .HTMLBody = "<p style=""font-family:Calibri (Body); font-size:11.5pt;"">" & Greeting & firstname & "," & r1 & strTo & "</p>" & .HTMLBody 'strbody & "<br>" & .HTMLBody
    .Display
    End If
End With

    
End Sub

