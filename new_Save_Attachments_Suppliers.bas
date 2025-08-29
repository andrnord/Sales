Attribute VB_Name = "new_Save_Attachments_Suppliers"
Sub Save_Attachments_Suppliers()

    Dim quotationNumber As String
    Dim searchRegExp As Object
    Dim match As Object
    Dim oMail As Outlook.MailItem
    Dim olAttachment As Attachment
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection

    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    Set oMail = myOlSel.item(1)

    ' Step 1: Extract quotation number using pattern SA6xxxx / SA7xxxx
    Set searchRegExp = CreateObject("VBScript.RegExp")
    searchRegExp.Pattern = "\bSA[67]\d{4}\b"
    searchRegExp.IgnoreCase = True
    searchRegExp.Global = False

    quotationNumber = ""

    If searchRegExp.test(oMail.Subject) Then
        Set match = searchRegExp.Execute(oMail.Subject)
        quotationNumber = match(0).Value
    ElseIf searchRegExp.test(oMail.Body) Then
        Set match = searchRegExp.Execute(oMail.Body)
        quotationNumber = match(0).Value
    End If

    If quotationNumber = "" Then quotationNumber = "SA70"

    quotationNumber = InputBox("Enter Quotation number (e.g., SA71234)...", "Quotation number", quotationNumber)
    If StrPtr(quotationNumber) = 0 Then Exit Sub

    If MsgBox("Are you sure Quotation number:" & vbCrLf & Space(15) & quotationNumber & "  is correct?", vbYesNo, "Continue?") = vbNo Then Exit Sub

    ' Step 2: Locate destination folder
    Dim rootPath As String
    rootPath = "C:\Users\andreas.nordvik\OneDrive - Transmark Subsea AS\Business Central - Salgsmulighet\"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderPath As String
    folderPath = FindFolderRecursive(rootPath, quotationNumber, fso)
    If folderPath = "" Then
        MsgBox "Quotation folder not found. Please check the number or folder location.", vbExclamation
        Exit Sub
    End If

    Dim FilePath As String
    FilePath = folderPath & "\03 Underleverandører\"
    Checking_If_Folder_Exists FilePath

    ' Step 3: Save attachments
    For Each oMail In myOlSel
        If TypeName(oMail) = "MailItem" And oMail.Attachments.Count > 0 Then
            For Each olAttachment In oMail.Attachments
                If LCase(Left(olAttachment.FileName, 5)) <> "image" Then
                    olAttachment.SaveAsFile fso.BuildPath(FilePath, olAttachment.FileName)
                End If
            Next olAttachment
        End If
    Next oMail

    Set fso = Nothing

End Sub

' Recursively finds folder containing quotation number
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

' Create folder if it does not exist
Sub Checking_If_Folder_Exists(FilePath As String)
    If Dir(FilePath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir FilePath
        On Error GoTo 0
    End If
End Sub

