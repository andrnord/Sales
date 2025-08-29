Attribute VB_Name = "AutoGreetingReply"
Sub AutoAddGreetingtoReply()
    Dim oMail As MailItem
    Dim oReply As MailItem
    Dim Sender As String, emailAdress As String, eFirst As String, eLast As String
    Dim firstname As String, lastname As String, tempname As String
    Dim Greeting As String
    Dim nameArray As Variant, mailArray As Variant
    Dim Recipient As String
     
    Select Case Application.ActiveWindow.Class
           Case olInspector
                Set oMail = ActiveInspector.CurrentItem
           Case olExplorer
                Set oMail = ActiveExplorer.Selection.item(1)
    End Select
    
    Greeting = "Hi "
    Sender = oMail.Sender
    Recipient = oMail.Recipients(1)
    
    If Sender = "Andreas Nordvik" Then
        Sender = Recipient
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
        Greeting = "Hei "
        
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
    
    Dim formattedGreeting As String
    formattedGreeting = "<html><body><p style='font-family:Aptos; font-size:12pt;'>" & Greeting & firstname & ",</p></body></html>"

With oReply
    .HTMLBody = formattedGreeting & .HTMLBody
    .Display
End With
    
End Sub

