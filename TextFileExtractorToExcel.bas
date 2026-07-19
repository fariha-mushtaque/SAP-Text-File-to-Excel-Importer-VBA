Attribute VB_Name = "Module1"
Option Explicit

Sub ImportTextFilesToSheets()

    Dim FolderPath As String
    Dim FileName As String
    Dim FileList() As String
    Dim FileCount As Long
    Dim i As Long

    Dim ws As Worksheet
    Dim qt As QueryTable

    Dim fd As FileDialog

    Dim SheetName As String

    Dim Imported As Long
    Dim Skipped As Long
    Dim Failed As Long

    Dim FailedFiles As String

    Dim Response As VbMsgBoxResult

    Dim StartTime As Double
    Dim EndTime As Double

    StartTime = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

'====================================================
'Select Folder
'====================================================

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd

        .Title = "Select Folder Containing Text Files"

        If .Show <> -1 Then
            MsgBox "Operation Cancelled.", vbInformation
            GoTo ExitMacro
        End If

        FolderPath = .SelectedItems(1) & "\"

    End With

'====================================================
'Collect TXT Files
'====================================================

    FileName = Dir(FolderPath & "*.txt")

    Do While FileName <> ""

        If Left(FileName, 2) <> "~$" Then

            FileCount = FileCount + 1
            ReDim Preserve FileList(1 To FileCount)
            FileList(FileCount) = FileName

        End If

        FileName = Dir

    Loop

    If FileCount = 0 Then

        MsgBox "No text files found in selected folder.", vbExclamation
        GoTo ExitMacro

    End If

'====================================================
'Sort Alphabetically
'====================================================

    Dim Temp As String
    Dim j As Long

    For i = 1 To FileCount - 1

        For j = i + 1 To FileCount

            If UCase(FileList(i)) > UCase(FileList(j)) Then

                Temp = FileList(i)
                FileList(i) = FileList(j)
                FileList(j) = Temp

            End If

        Next j

    Next i

'====================================================
'Import Files
'====================================================

    For i = 1 To FileCount

        FileName = FileList(i)

        On Error GoTo ImportError

        SheetName = Left(FileName, InStrRev(FileName, ".") - 1)

        SheetName = Replace(SheetName, "\", "_")
        SheetName = Replace(SheetName, "/", "_")
        SheetName = Replace(SheetName, "*", "_")
        SheetName = Replace(SheetName, "[", "_")
        SheetName = Replace(SheetName, "]", "_")
        SheetName = Replace(SheetName, ":", "_")
        SheetName = Replace(SheetName, "?", "_")

        If Len(SheetName) > 31 Then
            SheetName = Left(SheetName, 31)
        End If

'====================================================
'Duplicate Sheet
'====================================================

        If WorksheetExists(SheetName) Then

            Response = MsgBox( _
            "Worksheet '" & SheetName & "' already exists." & vbCrLf & vbCrLf & _
            "Yes = Replace" & vbCrLf & _
            "No = Skip" & vbCrLf & _
            "Cancel = Stop Import", _
            vbYesNoCancel + vbQuestion)

            Select Case Response

                Case vbYes

                    Worksheets(SheetName).Delete

                Case vbNo

                    Skipped = Skipped + 1
                    GoTo NextFile

                Case vbCancel

                    GoTo ExitMacro

            End Select

        End If

'====================================================
'Create Sheet
'====================================================

        Set ws = Worksheets.Add(After:=Sheets(Sheets.Count))

        ws.Name = SheetName

        Set qt = ws.QueryTables.Add( _
                Connection:="TEXT;" & FolderPath & FileName, _
                Destination:=ws.Range("A1"))

        With qt

            .TextFileParseType = xlDelimited
            .TextFileTabDelimiter = True
            .TextFileCommaDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileSpaceDelimiter = False

            .AdjustColumnWidth = True

            .Refresh BackgroundQuery:=False

            .Delete

        End With

        ws.Columns.AutoFit

        Imported = Imported + 1

NextFile:

        On Error GoTo 0

    Next i

    GoTo ExitMacro

'====================================================
'Error Handling
'====================================================

ImportError:

    Failed = Failed + 1

    FailedFiles = FailedFiles & vbCrLf & FileName

    Resume NextFile

'====================================================
'Finish
'====================================================

ExitMacro:

    EndTime = Timer

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim Msg As String

    Msg = "Import Completed Successfully!" & vbCrLf & vbCrLf

    Msg = Msg & "Folder :" & vbCrLf
    Msg = Msg & FolderPath & vbCrLf & vbCrLf

    Msg = Msg & "Total Files : " & FileCount & vbCrLf
    Msg = Msg & "Imported : " & Imported & vbCrLf
    Msg = Msg & "Skipped : " & Skipped & vbCrLf
    Msg = Msg & "Failed : " & Failed & vbCrLf & vbCrLf

    Msg = Msg & "Execution Time : " & _
          Format(EndTime - StartTime, "0.00") & " seconds"

    If Failed > 0 Then

        Msg = Msg & vbCrLf & vbCrLf & _
              "Failed Files:" & vbCrLf & FailedFiles

    End If

    MsgBox Msg, vbInformation

End Sub

Function WorksheetExists(SheetName As String) As Boolean

    On Error Resume Next

    WorksheetExists = Not Worksheets(SheetName) Is Nothing

    On Error GoTo 0

End Function

