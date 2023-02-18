Attribute VB_Name = "Update"


Sub download_file(myURL As String, myPath As String, FileName As String)
Dim oStrm As Object
Dim HttpReq As Object
Set HttpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
HttpReq.Open "GET", myURL, False, "username", "password"
HttpReq.Send

myURL = HttpReq.responseBody
If HttpReq.Status = 200 Then
    Set oStrm = CreateObject("ADODB.Stream")
    oStrm.Open
    oStrm.Type = 1
    oStrm.Write HttpReq.responseBody
    oStrm.SaveToFile ThisWorkbook.Path & "\" & FileName, 1 ' 1 = no overwrite, 2 = overwrite
    oStrm.Close
End If
End Sub

Sub Update_Check()

'Refres cell and check equalization
'if true do nothing
'else take current file name and current file path
'download new file to this path
'change current filename as temp and new file name to current
'open new file
'delete old file


Dim CurrentVersion, NewVersion As String

Dim DownloadLink As String

Dim result As String

Dim myPath As String

Dim oldFileName As String

Dim wkb As Workbook
Dim a As Integer
Dim un As String
a = 5

 myPath = ThisWorkbook.Path

 oldFileName = ThisWorkbook.Name
un = "Dear " & Environ("UserName")
On Error GoTo baglanti_hatasi

ThisWorkbook.Worksheets("Data Cloud").ListObjects("Data").QueryTable.Refresh BackgroundQuery:=False

CurrentVersion = ThisWorkbook.Worksheets("Data Local").Range("B2").Value

NewVersion = ThisWorkbook.Worksheets("Data Cloud").Range("B2").Value

DownloadLink = ThisWorkbook.Worksheets("Data Cloud").Range("B3").Value


        If CurrentVersion <> NewVersion Then
            
           result = MsgBox("There is a Newer Version of This File" & vbNewLine & "Click Yes to Update", vbYesNo + vbQuestion, "Update File")
            
                If result = vbYes Then
                    On Error GoTo dlhatasi
                    
                    ActiveWorkbook.SaveAs Environ("Temp") & "\Temp786954.xlsm"
                    Kill myPath & "\" & oldFileName
                    Call download_file(DownloadLink, myPath, oldFileName)
                    
                    Workbooks.Open (myPath & "\" & oldFileName)
                    
                     
                 
                
                Else
                
                
                End If

            
        Else
        
            Call subClosingPopUp(1, "The File Is Up To Date", un, 1)
            
            
        
        End If
        
    If IsWorkBookOpen(Environ("Temp") & "\Temp786954.xlsm") Then
            
    
            Application.Workbooks("Temp786954.xlsm").Close SaveChanges:=False
            
            
    End If
    
'my way of err hendling :)
    If a = 21 Then

baglanti_hatasi:

    msg1 = MsgBox("Data Source is Not Available" & vbNewLine & "Please Try Again Later", vbCritical, un)
    Exit Sub

    End If
    
If a = 44 Then
dlhatasi:

msg = MsgBox("Could Not Download The File!", vbExclamation, un)
ActiveWorkbook.SaveAs myPath & oldFileName
End If

End Sub

Sub KillTempFile()

Dim myPath As String
Dim strFileName As String
Dim strFileExists As String

myPath = Environ("Temp")

strFileName = myPath & "\Temp786954.xlsm"


Do
    
    strFileExists = Dir(strFileName)
    If strFileExists <> "" Then
    
    
        If IsWorkBookOpen(strFileName) Then
            
            Workbooks("Temp786954.xlsm").Close SaveChanges:=True
        
         Else
    
        Kill myPath & "\Temp786954.xlsm"
        End If
    End If
Loop While strFileExists <> ""

myPath = ThisWorkbook.Path

strFileName = myPath & "\Temp786954.xlsm"

Do
    
    strFileExists = Dir(strFileName)
    If strFileExists <> "" Then
    
    
        If IsWorkBookOpen(strFileName) Then
            
            Workbooks("Temp786954.xlsm").Close SaveChanges:=True
        
         Else
    
        Kill myPath & "\Temp786954.xlsm"
        End If
    End If
Loop While strFileExists <> ""
    


End Sub

Sub refresh_data()

ActiveWorkbook.RefreshAll



End Sub

Function IsWorkBookOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case 53: Exit Function
    
    Case Else: Error ErrNo
    End Select
End Function


Sub SuicideSub()
' Original code from Tom Urtis

'MsgBox "I'm gonna do it!!!"



With ThisWorkbook

.Saved = True

.ChangeFileAccess xlReadOnly

Kill .FullName

Application.DisplayAlerts = False

ThisWorkbook.Close False

'Application.Quit

End With

End Sub
