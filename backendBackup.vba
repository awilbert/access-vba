' Create a backup of a back-end database from the front-end.
' Script can be run from a button or other triggering event.
' It will determine the name and current path of the back-end file, and
' copy a weekly file into a "backup" subfolder. 
' Backups tagged with the day of the week, and will overwrite old backups after 7 days.
' Script will fail gracefully if the back-end is locked. Make sure all tables/forms/queries are closed before running.


Private Sub btnBackup_Click()

On Error GoTo Err_backup
    
Dim strFullPath As String
Dim strBackendFile As String
Dim strPath As String
Dim strSourceFile As String
Dim strDestinationFile As String

' get path to back-end using a linked table as reference. Change "tblLinked" to the name of a table within the back-end db.
' Mid function drops connection info including password (starts at character position 01)
' the 01 will need to be customized depending on the existence and length of the back-end database's password.

strFullPath = Mid(DBEngine.Workspaces(0).Databases(0).TableDefs("tblLinked").Connect, 01)

' uncomment next line in order to check path string in order to fine-tune the number of characters to 
' truncate form the begining (by changing #01 above) in order to arrive at just the file path.

'MsgBox (strFullPath)

' isolate the name of the backend database
    For I = Len(strFullPath) To 1 Step -1
        If Mid(strFullPath, I, 1) = "\" Then
            strBackendFile = Mid(strFullPath, (I + 1))
            Exit For
        End If
    Next

    For N = Len(strBackendFile) To 1 Step -1
        If Mid(strBackendFile, N, 1) = "." Then
            strBackendFile = Left(strBackendFile, (N - 1))
            Exit For
        End If
    Next

' remove the filename of the database to isolate the path
    For I = Len(strFullPath) To 1 Step -1
        If Mid(strFullPath, I, 1) = "\" Then
            strPath = Left(strFullPath, I)
            Exit For
        End If
    Next

' reconstruct the source and destination file paths
' destination file to include the abbreviated day of the week
' \backup directory must already exist

    strSourceFile = strPath & strBackendFile & ".accdb"
    strDestinationFile = strPath & "backup\" & strBackendFile & "-" & WeekdayName(Weekday(Date), True) & ".accdb"

FileCopy strSourceFile, strDestinationFile
MsgBox "The back-end database has been backed up!"

Exit_Backup:
Exit Sub

Err_backup:
If Err.Number = 0 Then
    ElseIf Err.Number = 70 Then
        MsgBox "The file is currently in use and therefore is locked and cannot be copied at this time. Please ensure that all forms, reports, and queries are closed, and that no one is using the database and try again.", vbOKOnly, "File Currently in Use"
    ElseIf Err.Number = 53 Then
        MsgBox "The Source File '" & strSourceFile & "' could not be found. Please validate the location and name of the specifed Source File and try again", vbOKOnly, "File Not Found"
    Else
        MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Error Source: ModExtFiles / CopyFile" & vbCrLf & "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
End If

Resume Exit_Backup
    
End Sub

