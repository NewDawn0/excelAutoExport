Attribute VB_NAME = "Config"
' This defines the name of the module as "Config" '

Type export
  startCell As String ' Description: The starting cell of the range to be exported from Excel. '
                      ' Example: "A2" specifies the cell A2 in the current worksheet. '
  endCell As String   ' Description: The ending cell of the range to be exported. '
                      ' Example: "B4" specifies the cell B4 in the current worksheet. '
  marker As String    ' Description: The placeholder text (marker) in the Word document where the exported '
                      ' data will be pasted. This marker cannot include special characters like "!" '
                      ' Example: "MyMarker" is a placeholder in the Word document that will be replaced with the data. '
  file As String      ' Description: The full path to the Word template file (.dotm) '
                      ' that contains the marker where the exported data will go. '
                      ' Example: "C:\Templates\ReportTemplate.dotm" '
End Type

' Description: This function creates and returns an array of export configurations. '
' Example: Use this function to create an export '
Function config() As export()
  ' Description: Declare an array to hold the export configurations. '
  ' Example: For one export, use (1 To 1). For two exports, use (1 To 2), and so on. '
  Dim exports(1 To 1) As export
  
  exports(1).startCell = "A1"     ' Set starting cell to A1 '
  exports(1).endCell = "C2"       ' Set ending cell range to C2 '
  exports(1).marker = "marker"    ' Bookmark it should find in the exports(n).file '
  exports(1).file = "./test.docx" ' File to export the data to '
                                  ' Can be absolute or relative path '
                                  ' Example: C:\Users\ACIVI\file.docx '
                                  ' Example: .\file.docx '

  config = exports ' Returns the array of exports '
End Function