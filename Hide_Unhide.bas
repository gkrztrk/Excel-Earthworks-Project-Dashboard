Attribute VB_Name = "Hide_Unhide"
Option Explicit

Sub Hide_Frames()
On Error Resume Next
Application.DisplayFullScreen = True
Application.DisplayFormulaBar = False
ActiveWindow.DisplayWorkbookTabs = True
ActiveWindow.DisplayHeadings = False
ActiveWindow.DisplayGridlines = False


End Sub

Sub Unhide_Frames()
On Error Resume Next
Application.DisplayFullScreen = False
Application.DisplayFormulaBar = True
ActiveWindow.DisplayWorkbookTabs = True
ActiveWindow.DisplayHeadings = True
ActiveWindow.DisplayGridlines = True


End Sub

Sub SetSpecificScrollArea()
Dim ws As Worksheet
Set ws = ActiveSheet
ws.ScrollArea = "A1:AT50"

End Sub

Sub fullscreen()

If Application.DisplayFullScreen = True Then
    On Error Resume Next
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayGridlines = True
    
Else
    On Error Resume Next
    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayGridlines = False

End If

End Sub
