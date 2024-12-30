Attribute VB_Name = "Main"
' Sets the module name to "Main". '

' Main procedure to export data to Word documents. '
Public Sub ExportData()
  Dim resp As VbMsgBoxResult ' Variable to store the user's response to a message box. '
  
  ' Display a warning message to the user. '
  ' Description: Informs the user that the macro will close all open Word documents and prompts them to confirm. '
  resp = MsgBox("This macro will close all your Word documents. " & vbCrLf & _
                "Make sure all your Word documents are saved.", _
                vbOKCancel + vbInformation, "Do you want to proceed?")
  
  ' Check the user's response to the message box. '
  If resp = vbCancel Then
    End ' Exit the macro if the user chooses "Cancel". '
  End If

  ' Declare an array to hold export configurations. '
  Dim exports() As config.export
  
  ' Retrieve the export configurations defined in the "config" module. '
  exports = config.config()
  
  ' Call the utility function to perform the export. '
  ' Description: Passes the export configurations to the ExportToWord function, '
  ' which handles copying data to Word based on the configurations. '
  Util.ExportToWord exports
  
  ' Notify the user that the export was successful. '
  MsgBox "Successfully exported all data", vbInformation, "Export"
End Sub
