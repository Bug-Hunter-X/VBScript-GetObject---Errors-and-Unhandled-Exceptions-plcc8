Function GetObject() is used to create an object, but if the object doesn't exist, it throws an error.  This can happen if the object is not properly registered, or if the path to the object is incorrect.  The error is not always informative, making debugging difficult.  Example:

```vbscript
Set objExcel = GetObject(, "Excel.Application")
' Error if Excel is not running or not installed
```

Another example is improper handling of errors from external libraries or COM objects.  A simple example using `Err.Number` which may not capture all error situations:

```vbscript
On Error Resume Next
Set objFSO = CreateObject("Scripting.FileSystemObject")
if Err.Number <> 0 then
  WScript.Echo "Error creating FileSystemObject: " & Err.Description
end if
Set objFSO = Nothing
```