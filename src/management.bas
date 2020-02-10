Attribute VB_Name = "management"
' start nano and create necessary sheet for authentication
Private Sub OpenNano()

    On Error GoTo Err
    If Not (CreateAuthSheet(CStr(Range("currentNano").Value), CStr(Range("user").Value))) Then
        GoTo Err
    End If
    
    If Not (CreateNano(CStr(Range("currentNano").Value))) Then
        GoTo Err
    End If
    Exit Sub
    
Err:

    Application.Run ("PageSetup.CloseCleanup")
    Exit Sub
    
End Sub

' start nano from the server
Private Function CreateNano(label As String) As Boolean
    CreateNano = True

    Range("status").Value = "attaching nano"
    Dim Client As New WebClient
    
    On Error GoTo Err
    Client.BaseUrl = Worksheets(label).Range("url").Value
    Client.TimeoutMs = 75000
    
    Dim Request As New WebRequest
    Request.Resource = "nanoInstance/{label}"
    Request.Method = WebMethod.HttpPost
    Request.AddUrlSegment "label", label
    Request.AddQuerystringParam "api-tenant", Worksheets(label).Range("apitenant").Value
    Request.AddHeader "x-token", Worksheets(label).Range("xtoken").Value

    Dim Response As WebResponse
    Set Response = Client.Execute(Request)
    
    On Error GoTo JSONErr
    Dim json As Object
    Set json = JsonConverter.ParseJson(Response.Content)
    If Response.StatusCode <> 200 Then
        MsgBox "NANO ERROR:" & vbNewLine & "   " & json("message")
        CreateNano = False
    Else
        Worksheets(label).Range("instance").Value = json("instanceID")
    End If
    
    Worksheets(label).Protect
    
    Range("status").Value = "finished"
    
Exit Function
    
Err:
    MsgBox "Instance error: reattach to instance"
    CreateNano = False
    Exit Function
    
JSONErr:
    MsgBox "Server error. Check instance is running"
    CreateNano = False
    Exit Function

End Function


Private Function ReadAuthFile() As String
    ' read authentication file
    Dim file As String, textline As String
    
    If InStr(Application.OperatingSystem, "Windows") > 0 Then
        ' windows
        file = "C:\Users\" & Environ("Username") & "\.BoonLogic.lic"
    Else
        ' macos (or linux???)
        file = "Macintosh HD:Users:" & MacScript("set userName to short user name of (system info)" & vbNewLine & "return userName") & ":.BoonLogic.lic"
    End If
    
    On Error GoTo Err:
    Open file For Input As #1
    Do Until EOF(1)
        Line Input #1, textline
        ReadAuthFile = ReadAuthFile & textline
    Loop
    Close #1
    Exit Function
    
Err:
    MsgBox "Cannot find .Boonlogic file"
    ReadAuthFile = "False"
    Exit Function
    
End Function

Private Function GetUsers() As String
    Dim json As Object, Text As String
    Text = ReadAuthFile()
    GetUsers = "default"
    If Text = "False" Then
        GetUsers = "False"
        Exit Function
    End If
    Set json = JsonConverter.ParseJson(Text)
    For Each Key In json.Keys
        If Key <> "default" Then
            GetUsers = GetUsers & "," & Key
        End If
    Next

End Function

' creates sheet that holds authentication information for that nano
Private Function CreateAuthSheet(label As String, user As String) As Boolean
    CreateAuthSheet = True
    
    Dim ws As Worksheet
    If WorksheetExists(label) Then
        Application.DisplayAlerts = False
        Worksheets(label).Delete
        Application.DisplayAlerts = True
    End If
    
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = label
    ws.Visible = xlSheetHidden
 
    ws.Cells(1, 1).Name = "xtoken"
    ws.Cells(2, 1).Name = "url"
    ws.Cells(3, 1).Name = "apitenant"
    ws.Cells(4, 1).Name = "instance"

    Dim Text As String
    Text = ReadAuthFile()
    
    On Error GoTo Err
    Dim json As Object
    Set json = JsonConverter.ParseJson(Text)
    ws.Range("xtoken").Value = json(user)("api-key")

    ws.Range("url").Value = json(user)("server") & "/expert/v3/"
    
    ws.Range("apitenant").Value = json(user)("api-tenant")
    
    Worksheets("BoonNano").Activate
    
Exit Function

Err:
    MsgBox "User not found"
    CreateAuthSheet = False
    Exit Function

End Function

' stop instance and delete sheets corresponding to it
Private Function CloseNano() As Boolean
    CloseNano = True

    Range("status").Value = "closing nano"
    Dim label As String
    label = Range("currentNano").Value
    
    Dim Client As New WebClient
    
    On Error GoTo Err
    Client.BaseUrl = Worksheets(label).Range("url").Value
    Client.TimeoutMs = 75000

    Dim Request As New WebRequest
    Request.Resource = "nanoInstance/{label}"
    Request.Method = WebMethod.HttpDelete
    Request.AddUrlSegment "label", label
    Request.AddQuerystringParam "api-tenant", Worksheets(label).Range("apitenant").Value
    Request.AddHeader "x-token", Worksheets(label).Range("xtoken").Value

    Dim Response As WebResponse
    Set Response = Client.Execute(Request)
    
    On Error GoTo JSONErr
    Dim json As Object
    Set json = JsonConverter.ParseJson(Response.Content)
    If Response.StatusCode <> 200 Then
        MsgBox "NANO ERROR:" & vbNewLine & "   " & json("message")
        CloseNano = False
    End If
    
    If WorksheetExists(label) Then
        Application.DisplayAlerts = False
        Worksheets(label).Delete
        Application.DisplayAlerts = True
    End If
    
    Range("status").Value = "finished"
    
Exit Function

Err:
    CloseNano = False
    Exit Function

JSONErr:
    CloseNano = False
    Exit Function
    
End Function

' check if the worksheet exists
Private Function WorksheetExists(ByVal WorksheetName As String) As Boolean
    Dim Sht As Worksheet

      For Each Sht In ActiveWorkbook.Worksheets
           If Application.Proper(Sht.Name) = Application.Proper(WorksheetName) Then
               WorksheetExists = True
               Exit Function
           End If
        Next Sht
    WorksheetExists = False
End Function

