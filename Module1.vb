Option Explicit

Sub test2()
    Debug.Print "New line"
    Dim Index As Integer
    For Index = 1 To 500
        Debug.Print Error$(Index)
    Next Index
End Sub