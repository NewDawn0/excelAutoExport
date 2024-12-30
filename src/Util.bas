Attribute VB_NAME = "Util"
' Sets the module name to "Util". '

' Main subroutine to handle exporting data to Word documents. '
Sub ExportToWord(ByRef exports() As config.export)
  ' Validate the export configurations and perform the export. '
  Dim i As Long ' Loop counter for iterating through export configurations. '
  Dim filePath As String ' Variable to store the resolved file path. '
  Dim dataRange As Range ' Placeholder for a data range (not used in this implementation). '
  Dim bookmarkExists As Boolean ' Tracks if the bookmark exists (not directly used here). '
  bookmarkExists = True ' Initialize bookmarkExists. '

  ' Check prerequisites for each export configuration. '
  For i = LBound(exports) To UBound(exports)
    ' Resolve file path: handle relative paths (starting with "./" or ".\") or absolute paths. '
    filePath = exports(i).file
    If Left(filePath, 2) = "./" Then
      resolvedPath = ActiveWorkbook.Path & "/" & Mid(filePath, 3)
    ElseIf Left(filePath, 2) = ".\" Then
      resolvedPath = ActiveWorkbook.Path & "\" & Mid(filePath, 3)
    Else
      resolvedPath = filePath
    End If

    ' Check if the file exists. If not, show an error and abort. '
    If Dir(resolvedPath) = "" Then
      MsgBox "The file '" & resolvedPath & "' does not exist.", vbCritical, "File Not Found"
      End ' Abort if the file is missing. '
    End If

    ' Update the file path in the export configuration. '
    exports(i).file = resolvedPath

    ' Check if the marker exists in the Word document. '
    If checkMarker(exports(i)) = False Then
      MsgBox "The marker '" & exports(i).marker & "' does not exist in file " & exports(i).file & ".", vbCritical, "Marker Not Found"
      End ' Abort if the marker is missing. '
    End If
  Next i

  ' Perform data export for each configuration. '
  For i = LBound(exports) To UBound(exports)
    cpyData exports(i)
  Next i
End Sub

' Checks if the specified marker (bookmark) exists in the Word document. '
Function checkMarker(ByRef export As config.export) As Boolean
  Dim wdApp As Word.Application ' Word application object. '
  Dim wdDoc As Word.Document ' Word document object. '
  Dim Res As Boolean ' Tracks whether the marker exists. '
  Res = False ' Initialize result. '

  ' Start a new instance of Word. '
  Set wdApp = New Word.Application
  With wdApp
    .Visible = True ' Make Word visible. '
    .Activate ' Bring Word to the foreground. '
  End With

  ' Open the Word document and check for the marker. '
  Set wdDoc = wdApp.Documents.Open(export.file)
  With wdDoc
    Res = .Bookmarks.Exists(export.marker) ' Check if the bookmark exists. '
    .Close ' Close the document. '
  End With

  ' Quit Word and release resources. '
  wdApp.Quit
  Set wdDoc = Nothing
  Set wdApp = Nothing

  ' Return the result (True if marker exists, False otherwise). '
  checkMarker = Res
End Function

' Copies data from Excel and pastes it into the specified Word document at the specified marker. '
Function cpyData(ByRef export As config.export)
  Dim wdApp As Word.Application ' Word application object. '
  Dim wdDoc As Word.Document ' Word document object. '

  ' Start a new instance of Word. '
  Set wdApp = New Word.Application
  With wdApp
    .Visible = True ' Make Word visible. '
    .Activate ' Bring Word to the foreground. '
  End With

  ' Open the Word document. '
  Set wdDoc = wdApp.Documents.Open(export.file)

  ' Copy data from Excel. '
  ActiveSheet.Range(export.startCell, export.endCell).Copy
  
  ' Pause for 1 second to allow clipboard operations to complete. '
  Application.Wait Now() + #12:00:01 AM#

  ' Paste the copied data into the Word document at the specified marker. '
  With wdDoc
    .Bookmarks(export.marker).Range.PasteExcelTable False, True, False ' Paste as a table. '
    .Save ' Save the document. '
    .Close ' Close the document. '
  End With

  ' Quit Word and release resources. '
  wdApp.Quit
  Set wdDoc = Nothing
  Set wdApp = Nothing

  ' Clear the clipboard to free up memory. '
  Application.CutCopyMode = False
End Function