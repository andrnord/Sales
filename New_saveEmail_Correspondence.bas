Attribute VB_Name = "New_saveEmail_Correspondence"
Option Explicit
Dim numberOfEmails As Double

Sub saveEmailCorrespondence()

    Dim X As Long
    Dim quotationNumber As String
    Dim searchRegExp As Object
    Dim match As Object
    Dim oMail As Outlook.MailItem
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection

    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    Set oMail = myOlSel.item(1)

    Dim sSubject As String
    Dim sTime As String
    Dim sDate As String
    Dim iYear As Integer
    Dim sBody As String

    ' Extracting date/time format
    sTime = Replace(oMail.ReceivedTime, ":", ".")
    sDate = Split(sTime, " ")(0)
    iYear = Right(sDate, 4)
    sTime = iYear & Split(sTime, ".")(1) & Split(sTime, ".")(0) & "_" & Split(sTime, " ")(1)
    sSubject = oMail.Subject
    sBody = oMail.Body

    ' Step 1: Search for quotation number
    Set searchRegExp = CreateObject("VBScript.RegExp")
    searchRegExp.Pattern = "\bSA[67]\d{4}\b"
    searchRegExp.IgnoreCase = True
    searchRegExp.Global = False
    quotationNumber = ""

    If searchRegExp.test(sSubject) Then
        Set match = searchRegExp.Execute(sSubject)
        quotationNumber = match(0).Value
    ElseIf searchRegExp.test(sBody) Then
        Set match = searchRegExp.Execute(sBody)
        quotationNumber = match(0).Value
    End If

    If quotationNumber = "" Then quotationNumber = "SA70"

    quotationNumber = InputBox("Enter Quotation number (e.g., SA71234)...", "Quotation number", quotationNumber)
    If StrPtr(quotationNumber) = 0 Then Exit Sub

    If MsgBox("Are you sure Quotation number:" & vbCrLf & Space(15) & quotationNumber & "  is correct?", vbYesNo, "Continue?") = vbNo Then Exit Sub

    ' Step 2: Recursive search for quotation folder
    Dim rootPath As String
    rootPath = "C:\Users\andreas.nordvik\OneDrive - Transmark Subsea AS\Business Central - Salgsmulighet\"
    
    Dim fso As Object
    Dim folderPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    folderPath = FindFolderRecursive(rootPath, quotationNumber, fso)
    If folderPath = "" Then
        MsgBox "Quotation folder not found. Please check the number or folder location.", vbExclamation
        Exit Sub
    End If

    ' Step 3: Save emails to folder
    Dim targetPath As String
    targetPath = folderPath & "\02 Kundedialog\"
    Checking_If_Folder_Exists targetPath

    For X = 1 To myOlSel.Count
        Set oMail = myOlSel.item(X)
        sSubject = oMail.Subject

        ' Remove special characters
        sSubject = Replace(sSubject, "RE: ", "")
        sSubject = Replace(sSubject, "[External] ", "")
        sSubject = Replace(sSubject, "[EXT] ", "")
        sSubject = Replace(sSubject, "SV: ", "")
        
        sSubject = Replace(sSubject, ":", "_")
        sSubject = Replace(sSubject, ",", "_")
        sSubject = Replace(sSubject, "/", "_")
        sSubject = Replace(sSubject, "\", "_")
        sSubject = Replace(sSubject, "?", "_")
        sSubject = Replace(sSubject, "*", "_")
        sSubject = Replace(sSubject, ">", "_")
        sSubject = Replace(sSubject, "<", "_")
        sSubject = Replace(sSubject, "|", "_")
        sSubject = Replace(sSubject, Chr(34), "")

        sTime = Replace(oMail.ReceivedTime, ":", ".")
        sDate = Split(sTime, " ")(0)
        iYear = Right(sDate, 4)
        sTime = iYear & Split(sTime, ".")(1) & Split(sTime, ".")(0) & "_" & Left(Split(sTime, " ")(1), 5)

        Dim fullFilePath As String
        If oMail.SenderEmailType = "SMTP" Then
            fullFilePath = targetPath & sTime & "_" & StrConv(Split(Split(oMail.SenderEmailAddress, "@")(1), ".")(0), vbProperCase) & "_" & sSubject & ".msg"
        Else
            fullFilePath = targetPath & sTime & "_" & oMail.Sender & "_" & sSubject & ".msg"
        End If

        ' ?? Check if path length exceeds 260 and trim subject if needed
        If Len(fullFilePath) > 235 Then ' Leave a little buffer for extension
            Dim allowedLen As Long
            allowedLen = 235 - Len(targetPath & sTime & "_") ' space for date/time and dash
            sSubject = Left(sSubject, allowedLen)
            If oMail.SenderEmailType = "SMTP" Then
                fullFilePath = targetPath & sTime & "_" & StrConv(Split(Split(oMail.SenderEmailAddress, "@")(1), ".")(0), vbProperCase) & "_" & sSubject & ".msg"
            Else
                fullFilePath = targetPath & sTime & "_" & oMail.Sender & "_" & sSubject & ".msg"
            End If
        End If

        If FileExists(fullFilePath) Then GoTo NextX

        On Error GoTo SaveError
        oMail.SaveAs fullFilePath, OlSaveAsType.olMSGUnicode

NextX:
    Next X
    Exit Sub

SaveError:
    MsgBox "An error occurred while saving email: " & Err.Description

End Sub

Function FindFolderRecursive(basePath As String, quoteNumber As String, fso As Object) As String
    Dim folder As Object, subfolder As Object
    On Error Resume Next
    Set folder = fso.GetFolder(basePath)
    On Error GoTo 0
    If folder Is Nothing Then Exit Function
    
    For Each subfolder In folder.subFolders
        If InStr(subfolder.Name, quoteNumber) > 0 Then
            FindFolderRecursive = subfolder.Path
            Exit Function
        Else
            FindFolderRecursive = FindFolderRecursive(subfolder.Path, quoteNumber, fso)
            If FindFolderRecursive <> "" Then Exit Function
        End If
    Next
End Function

Sub Checking_If_Folder_Exists(FilePath As String)
    If Dir(FilePath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir FilePath
        On Error GoTo 0
    End If
End Sub

Function FileExists(FilePath As String) As Boolean
    FileExists = (Dir(FilePath) <> "")
End Function

