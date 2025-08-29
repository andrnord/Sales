Attribute VB_Name = "Handover"
Sub Handover()

    Dim quotationNumber As String
    Dim projectNumber As String
    Dim searchRegExp As Object
    Dim match As Object
    Dim rootQuotePath As String
    Dim rootProjectPath As String
    Dim quoteFolderPath As String
    Dim projectFolderPath As String
    Dim fso As Object
    Dim oMail As Object

    ' Step 1: Get Quotation Number
    Set searchRegExp = CreateObject("VBScript.RegExp")
    searchRegExp.Pattern = "\bSA[67]\d{4}\b"
    searchRegExp.IgnoreCase = True
    searchRegExp.Global = False

    quotationNumber = "SA70"
    If Application.ActiveExplorer.Selection.Count > 0 Then
        Set oMail = Application.ActiveExplorer.Selection(1)
        If searchRegExp.test(oMail.Subject) Then
            Set match = searchRegExp.Execute(oMail.Subject)
            quotationNumber = match(0).Value
        ElseIf searchRegExp.test(oMail.Body) Then
            Set match = searchRegExp.Execute(oMail.Body)
            quotationNumber = match(0).Value
        End If
    End If

    quotationNumber = InputBox("Enter Quotation number (e.g., SA71234)...", "Quotation number", quotationNumber)
    If StrPtr(quotationNumber) = 0 Then Exit Sub

    If MsgBox("Confirm quotation number:" & vbCrLf & quotationNumber, vbOKCancel, "Confirm Quotation") <> vbOK Then Exit Sub

    ' Step 2: Get Project Number (with automatic search in email)
    projectNumber = "J1"
    If Not oMail Is Nothing Then
        searchRegExp.Pattern = "\bJ1\d{4}\b"
        If searchRegExp.test(oMail.Subject) Then
            Set match = searchRegExp.Execute(oMail.Subject)
            projectNumber = match(0).Value
        ElseIf searchRegExp.test(oMail.Body) Then
            Set match = searchRegExp.Execute(oMail.Body)
            projectNumber = match(0).Value
        End If
    End If

    projectNumber = InputBox("Enter Project Number (e.g., J10135)...", "Project number", projectNumber)
    If StrPtr(projectNumber) = 0 Then Exit Sub

    If MsgBox("Confirm project number:" & vbCrLf & projectNumber, vbOKCancel, "Confirm Project") <> vbOK Then Exit Sub

    ' Step 3: Locate Quotation Folder
    rootQuotePath = "C:\Users\andreas.nordvik\OneDrive - Transmark Subsea AS\Business Central - Salgsmulighet\"
    Set fso = CreateObject("Scripting.FileSystemObject")

    quoteFolderPath = FindFolderRecursive(rootQuotePath, quotationNumber, fso)
    If quoteFolderPath = "" Then
        MsgBox "Quotation folder not found.", vbExclamation
        Exit Sub
    End If

    ' Step 4: Locate Project Folder
    rootProjectPath = "C:\Users\andreas.nordvik\OneDrive - Transmark Subsea AS\Business Central - Prosjekter\"
    projectFolderPath = FindFolderRecursive(rootProjectPath, projectNumber, fso)
    If projectFolderPath = "" Then
        MsgBox "Project folder not found.", vbExclamation
        Exit Sub
    End If

    ' Confirm Copy
    If MsgBox("Copy entire contents from:" & vbCrLf & quoteFolderPath & vbCrLf & "To existing project folder:" & vbCrLf & projectFolderPath, vbYesNo, "Confirm Copy") = vbNo Then
        Exit Sub
    End If

    ' Step 5: Copy Folder and All Contents
    Dim destinationPath As String
    destinationPath = projectFolderPath & "\" & fso.GetFolder(quoteFolderPath).Name

    If fso.FolderExists(destinationPath) Then
        MsgBox "Destination folder already exists. Handover aborted to prevent overwrite.", vbExclamation
        Exit Sub
    End If

    fso.CopyFolder quoteFolderPath, destinationPath

    MsgBox "Handover complete!", vbInformation

End Sub

' Recursively finds folder containing the target name
Function FindFolderRecursive(basePath As String, searchTerm As String, fso As Object) As String
    Dim folder As Object, subfolder As Object
    On Error Resume Next
    Set folder = fso.GetFolder(basePath)
    On Error GoTo 0
    If folder Is Nothing Then Exit Function

    For Each subfolder In folder.subFolders
        If InStr(1, subfolder.Name, searchTerm, vbTextCompare) > 0 Then
            FindFolderRecursive = subfolder.Path
            Exit Function
        Else
            FindFolderRecursive = FindFolderRecursive(subfolder.Path, searchTerm, fso)
            If FindFolderRecursive <> "" Then Exit Function
        End If
    Next
End Function

