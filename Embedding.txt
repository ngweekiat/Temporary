Sub RunEmbeddedPython()
    Dim scriptPath As String
    Dim pythonCode As String
    Dim f As Integer
    
    ' Read the embedded Python code from a cell
    pythonCode = Sheet1.Range("A1").Value

    ' Write it to a temporary .py file
    scriptPath = Environ("TEMP") & "\embedded_script.py"
    f = FreeFile
    Open scriptPath For Output As #f
    Print #f, pythonCode
    Close #f

    ' Run Python and keep the console open
    Shell "cmd /k python """ & scriptPath & """", vbNormalFocus
End Sub
