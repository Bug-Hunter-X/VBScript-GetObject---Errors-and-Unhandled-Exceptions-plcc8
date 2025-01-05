Improved error handling using more specific error checking and informative messages:

```vbscript
On Error GoTo ErrorHandler

Set objExcel = GetObject(, "Excel.Application")
' ... use objExcel ...

Exit Sub

ErrorHandler:
Select Case Err.Number
  Case 429  'ActiveX component can't create object
    WScript.Echo "Excel is not running or not installed. Please check your Excel installation."
  Case Else
    WScript.Echo "An error occurred: " & Err.Number & " - " & Err.Description
End Select
Err.Clear
```

For file system operations, use more specific error checks:

```vbscript
On Error GoTo FileError

Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FileExists("somefile.txt") Then
    WScript.Echo "File not found."
Else
    ' ... process the file ...
End If

Exit Sub

FileError:
  WScript.Echo "File system error: " & Err.Description
  Err.Clear
```