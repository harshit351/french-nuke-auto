Sub Main()
    Dim i As Long
    For i = 1 To 100
      'Your code here
      On Error GoTo ErrorHandler
        Call Function1
        Call Function2
      Application.StatusBar = "Processing step " & i & " of 100: " & Int(i / 100 * 100) & "% completed"
   Next i
   Application.StatusBar = False
    
    Exit Sub
 
ErrorHandler:
        MsgBox "Error # " & Err.Number & ": " & Err.Description
        Resume Next
        End Sub

Sub Function1()
    Application.StatusBar = "Running Function1"
    'Your code here
    Application.StatusBar = False
    End Sub
 
Sub Function2()
    Application.StatusBar = "Running Function2"
    'Your code here
    Application.StatusBar = False
     End Sub