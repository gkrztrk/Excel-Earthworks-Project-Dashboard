Attribute VB_Name = "Module1"
Option Explicit

Sub clearslicerfilter()

Dim objSlicer As SlicerCache
For Each objSlicer In ActiveWorkbook.SlicerCaches
     'If objSlicer.Name <> "RB" Then objSlicer.ClearManualFilter
     objSlicer.ClearManualFilter
Next objSlicer


End Sub

Sub lockwb()

Dim pword As String

pword = "EW"

   

    If ThisWorkbook.ProtectWindows Or ThisWorkbook.ProtectStructure Then

         UserForm1.Show
        
    Else
    
        Worksheets("FINANCE TABLE").Visible = False
        Worksheets("FINANCE").Visible = False


        ThisWorkbook.Protect Password:=pword
        
    End If
    
    
End Sub



Sub subClosingPopUp(PauseTime As Integer, Message As String, Title As String, ButtonType As Integer)

Dim WScriptShell As Object
Dim ConfigString As String

Set WScriptShell = CreateObject("WScript.Shell")
ConfigString = "mshta.exe vbscript:close(CreateObject(""WScript.Shell"")." & _
               "Popup(""" & Message & """," & PauseTime & ",""" & Title & """," & ButtonType & "))"
    'Sabit  Deger   Açiklama
    'vbOK     1     OK button
    'vbCancel 2     Cancel button
    'vbAbort  3     Abort button
    'vbRetry  4     Retry button
    'vbIgnore 5     Ignore button
    'vbYes    6     Yes button
    'vbNo     7     No button


WScriptShell.Run ConfigString

End Sub


