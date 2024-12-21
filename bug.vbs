Late Binding: VBScript's late binding can lead to runtime errors if an object or method doesn't exist.  This is especially problematic when dealing with COM objects or external libraries where the object's availability isn't guaranteed.
```vbscript
Dim objExcel
Set objExcel = CreateObject("Excel.Application")
' ... use objExcel ...
Set objExcel = Nothing
```
If the Excel application isn't installed, the `CreateObject` call will fail, potentially causing the script to crash.  Early binding (declaring object types explicitly) can help mitigate this but may require more upfront work.

Implicit Type Conversions: VBScript's loose typing can lead to unexpected type conversions. The script might not throw errors immediately but produce incorrect results silently.
```vbscript
Dim a, b
a = "10"
b = 5
Dim c
c = a + b 'Implicit conversion of "10" to 10; c will be 15
```
Adding a string ("10") to an integer (5) works without an error message, but if you're not aware of this implicit conversion, you might encounter unexpected behavior.

Error Handling: VBScript has limited structured error handling compared to more modern languages.  The `On Error Resume Next` statement can suppress errors, making debugging significantly harder.  Proper error handling requires careful planning and testing.