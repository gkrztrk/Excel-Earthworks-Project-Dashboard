Attribute VB_Name = "Log"
Sub SendToGoogle()
 
'This Macro Requires Reference to "Microsoft XML, v6.0" (VBA Editor > Tools > References, find &amp; select from list)
 
Dim URL_First As String        'Assign the first part of URL to send the data
Dim URL_Last As String         'Assign the last part of URL where we will update the information
Dim Form_URL As String         'To store the Form URL after merging Beginning and End URL
 
Dim HeaderName As String       'Variable to store the header type i.e. Content-Type
Dim SendID As String           'To store the information required to send a particular information to Google form
 
'Variables to store user inputs from Excel UserForm
Dim UsrName As String
Dim CompName As String
Dim FlName As String
Dim MacAddress As String
Dim LocationDetails As String
Dim eMail As String
 
'Assign User inputs to variables
 
UsrName = Environ("username")
CompName = Environ("computername")
FlName = "Dashboard"
MacAddress = GetMACAddress
LocationDetails = GetLocationDetails
eMail = GetUsersOutlookEmail

'https://docs.google.com/forms/d/e/1FAIpQLSf8UgSMFTbOjUDFJhAxSCfZL4RBhI8EjJawEYAQnZ81x5bjiA/formResponse?ifq

'&entry.2006625714=NAME&entry.812953752=COMP&entry.1687321223=FILE&entry.1112859251=MAC&entry.662911263=LOCATION
 
 '&entry.2006625714=Name&entry.812953752=comp&entry.1687321223=file&entry.1112859251=mac&entry.662911263=location&entry.1887480516=email
 
 
'Variable to store what we need to send to server
 
Dim TicketInfo As MSXML2.ServerXMLHTTP60 'XML variable to send the information to server
 
'Content-Type is actually a header type which tells the client what the content type of the returned content actually is. Google recognizes this header type
 
HeaderName = "Content-Type"
 
'SendID  required to send a particular information to Google Form
SendID = "application/x-www-form-urlencoded; charset=utf-8"
 
'In actual link, we need to replace viewform? with formResponse?ifq&amp;
'need to find the “name” attributes for the text boxes and the value for them
'add at the end &amp;submit=Submit and use it, it must post all the data you specified in one step.
 
'formRespose is used to get the response from Google Form after submitting the details
'Submit - it is a command to submit the filled form
 
URL_First = "https://docs.google.com/forms/d/e/1FAIpQLSf8UgSMFTbOjUDFJhAxSCfZL4RBhI8EjJawEYAQnZ81x5bjiA/formResponse?ifq"
 
URL_Last = "&entry.2006625714=" & UsrName & "&entry.812953752=" & CompName & "&entry.1687321223=" & FlName & "&entry.1112859251=" & MacAddress & "&entry.662911263=" & LocationDetails & "&entry.1887480516=" & eMail & "&submit=Submit"
 
'Creating the Final URL
Form_URL = URL_First & URL_Last
 
Set TicketInfo = New ServerXMLHTTP60 'Setting the reference of new server xmlhttp 60
 
TicketInfo.Open "POST", Form_URL, False ' Posting the entire link
 
TicketInfo.setRequestHeader HeaderName, SendID 'Specifies the name of an HTTP header.
 
TicketInfo.Send 'Send all the information over google
 
'StatusText is provide the status of data submission. It will show OK if data will be successfully submitted
 
If TicketInfo.statusText = "OK" Then 'Check for successful send
 
  'Call Reset 'Call Reset procedure to reset form Excel Form after submitting the data
  'MsgBox "Thank you for submitting data!"
 
Else
  MsgBox "Please check your internet connection &amp; required details"
End If
 
End Sub


Function GetMACAddress() As String
    Dim sComputer As String
    Dim oWMIService As Object
    Dim cItems As Object
    Dim oItem As Object
    Dim myMacAddress As String
    
    sComputer = "."
    
    Set oWMIService = GetObject("winmgmts:\\" & sComputer & "\root\cimv2")
    
    Set cItems = oWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
        
    For Each oItem In cItems
        If Not IsNull(oItem.IPAddress) Then myMacAddress = oItem.MacAddress
        Exit For
    Next
    'it will return mac address in format MM:MM:MM:SS:SS:SS
    'MsgBox myMacAddress
    GetMACAddress = myMacAddress

End Function

Sub Access_Check()

'Refres cell and check equalization
'if true do nothing
'else take current file name and current file path
'download new file to this path
'change current filename as temp and new file name to current
'open new file
'delete old file

Dim Banlist As Variant
Dim lr As Integer
Dim thisPc As String


Dim un As String


thisPc = GetMACAddress

un = "Dear " & Environ("UserName")
On Error GoTo baglanti_hatasi

ThisWorkbook.Worksheets("Data Cloud").ListObjects("Data").QueryTable.Refresh BackgroundQuery:=False

lr = ThisWorkbook.Worksheets("Data Cloud").Cells(Rows.Count, 5).End(xlUp).Row

Banlist = ThisWorkbook.Worksheets("Data Cloud").Range("E1:E" & lr).Value

If IsInArray(thisPc, Banlist) Then

    MsgBox "YOU DONT HAVE ACCESS THIS FILE!" & Chr(10) & "Contact: gokerozturk@yandex.com", vbCritical
    SuicideSub
Else
    
    MsgBox "WELCOME  " & un
    
End If

'my way of err hendling :)
    If a = 21 Then

baglanti_hatasi:

    msg1 = MsgBox("Data Source is Not Available" & vbNewLine & "Please Try Again Later", vbCritical, un)
    Exit Sub

    End If

End Sub

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i, 1) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function


Public Function GetUsersOutlookEmail(Optional ByVal errorFallback As String = "") As String
On Error GoTo catch
    With CreateObject("outlook.application")
        
        GetUsersOutlookEmail = .GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Parent.Name
    End With
Exit Function
catch:
    GetUsersOutlookEmail = errorFallback
End Function
