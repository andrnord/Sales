Attribute VB_Name = "birthday"
Sub CreateBirthdayNotifications()
    Dim OutlookApp As Object
    Dim Calendar As Object
    Dim BirthdaySheet As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim Name As String
    Dim LastName As String
    Dim birthday As Date
    Dim Events As Object

    ' Set the worksheet containing the birthday list
    Set BirthdaySheet = ThisWorkbook.Sheets("Sheet1") ' Change Sheet1 to the name of your sheet

    ' Find the last row in the list
    LastRow = BirthdaySheet.Cells(BirthdaySheet.Rows.Count, "A").End(xlUp).Row

    ' Create an instance of Outlook
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    If OutlookApp Is Nothing Then
        MsgBox "Outlook is not available. Please ensure Outlook is installed and try again.", vbExclamation
        Exit Sub
    End If

     ' Get the namespace and default calendar
    Set Namespace = OutlookApp.GetNamespace("MAPI")
    Set DefaultCalendar = Namespace.GetDefaultFolder(9) ' 9 = olFolderCalendar

    ' Search for the "Bursdager" folder
    On Error Resume Next
    Set CalendarFolder = Namespace.Folders("andreas.nordvik@transmark-subsea.com").Folders("Kalender").Folders("Bursdagspåminner")
    On Error GoTo 0

    If CalendarFolder Is Nothing Then
        MsgBox "The 'Bursdager' calendar folder was not found. Please ensure it exists under your default calendar.", vbExclamation
        Exit Sub
    End If
    
    ' Loop through each row in the birthday list
    For i = 3 To LastRow ' Assuming row 3 contains headers
        Name = BirthdaySheet.Cells(i, 1).Value
        LastName = BirthdaySheet.Cells(i, 2).Value
        birthday = BirthdaySheet.Cells(i, 4).Value

        If Name <> "" And IsDate(birthday) Then
            ' Create a new calendar event
            Set Events = CalendarFolder.Items.Add(1) ' 1 = olAppointmentItem
            With Events
                .Subject = "Birthday Reminder: " & Name
                .Start = birthday + TimeValue("9:00:00") ' Set reminder time to 9 AM
                .AllDayEvent = True
                .ReminderSet = True
                '.ReminderMinutesBeforeStart = 1440 ' 1 day before
                .Save
            End With
        End If
    Next i

    MsgBox "Birthday notifications have been created in Outlook!", vbInformation
End Sub



