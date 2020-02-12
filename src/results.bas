Attribute VB_Name = "results"

Private Function GetStatus() As Boolean
    GetStatus = True
    
    Dim label As String
    label = Range("currentNano").Value
    
    Range("status").Value = "getting results"
    Dim Client As New WebClient
    On Error GoTo Err
    Client.BaseUrl = Worksheets(label).Range("url").Value
    Client.TimeoutMs = 75000
    
    Dim Request As New WebRequest
    Request.Resource = "nanoStatus/{label}"
    Request.Method = WebMethod.HttpGet
    Request.AddUrlSegment "label", label
    Request.AddQuerystringParam "api-tenant", Worksheets(label).Range("apitenant").Value
    Request.AddQuerystringParam "results", "numClusters,totalInferences,averageInferenceTime"
    Request.AddHeader "x-token", Worksheets(label).Range("xtoken").Value

    Dim Response As WebResponse
    Set Response = Client.Execute(Request)
    
    On Error GoTo JSONErr
    Dim json As Object
    Set json = JsonConverter.ParseJson(Response.Content)
    If Response.StatusCode <> 200 Then
        MsgBox "NANO ERROR:" & vbNewLine & "   " & json("message")
        GetStatus = False
    Else
        Range("numClusters").Value = json("numClusters") - 1
        Range("totalInferences").Value = json("totalInferences")
        Range("avgClusterTime").Value = json("averageInferenceTime")
    End If
    
    Range("status").Value = "finished"
    
Exit Function

Err:
    MsgBox "Status call failed: " & Err.Description
    GetStatus = False
    Exit Function

JSONErr:
    MsgBox "Response error: status"
    GetStatus = False
    Exit Function

End Function

Private Function GetResults() As Variant
    GetResults = True
    
    Dim label As String
    label = Range("currentNano").Value
    
    Range("status").Value = "getting results"
    Dim Client As New WebClient
    
    On Error GoTo Err
    Client.BaseUrl = Worksheets(label).Range("url").Value
    Client.TimeoutMs = 75000
    
    Dim Request As New WebRequest
    Request.Resource = "nanoResults/{label}"
    Request.Method = WebMethod.HttpGet
    Request.AddUrlSegment "label", label
    Request.AddQuerystringParam "api-tenant", Worksheets(label).Range("apitenant").Value
    Request.AddQuerystringParam "results", "ID,SI,RI,DI,FI"
    Request.AddHeader "x-token", Worksheets(label).Range("xtoken").Value

    Dim Response As WebResponse
    Set Response = Client.Execute(Request)
    
    Dim json As Object
    On Error GoTo JSONErr
    Set json = JsonConverter.ParseJson(Response.Content)
    If Response.StatusCode <> 200 Then
        
        MsgBox "NANO ERROR:" & vbNewLine & "   " & json("message")
        GetResults = False
    End If
    
    Set GetResults = json
    Range("status").Value = "finished"
    
Exit Function
    
Err:
    MsgBox "Results call failed: " & Err.Description
    GetResults = False
    Exit Function

JSONErr:
    MsgBox "Response error: results"
    GetResults = False
    Exit Function

End Function

Function GetBufferStatus() As Variant
    GetBufferStatus = True
    
    Dim label As String
    label = Range("currentNano").Value
    
    Range("status").Value = "getting buffer status"
    Dim Client As New WebClient
    
    On Error GoTo Err
    Client.BaseUrl = Worksheets(label).Range("url").Value
    Client.TimeoutMs = 75000
    
    Dim Request As New WebRequest
    Request.Resource = "bufferStatus/{label}"
    Request.Method = WebMethod.HttpGet
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
        GetBufferStatus = False
    Else
        GetBufferStatus = json("totalyBytesProcessed") ' Response.Content
'        With Worksheets("BoonNano")
'            .Range("byteWritten").Value = json("totalBytesWritten")
'            .Range("byteProcess").Value = json("totalBytesProcessed")
'            .Range("byteBuffer").Value = json("totalBytesInBuffer")
'        End With
    End If
    
    Range("status").Value = "finished"
    
Exit Function

Err:
    MsgBox "Buffer status failed: " & Err.Description
    GetBufferStatus = False
    Exit Function

JSONErr:
    MsgBox "Response error: buffer status"
    GetBufferStatus = False
    Exit Function

End Function

Private Function LoadData() As Boolean
    LoadData = True

    On Error GoTo Err
    ' If Range("numFeatures") <> Selection.Columns.Count Then
    '     MsgBox "Feature count doesn't match. Reconfigure or choose different vector length"
    '     LoadData = False
    '     Exit Function
    ' End If
    
    Dim label As String
    label = Range("currentNano").Value
    
    Range("status").Value = "loading data"
    ' create selection as dictionary
    Dim row As Long, col As Long, arrString As String, tmpStr As String
    row = Selection.Rows.Count
    col = Selection.Columns.Count
    
    If InStr(Application.OperatingSystem, "Windows") > 0 Then
        returnStr = vbNewLine
    Else
        ' Macos or (linux??)
        returnStr = vbCr & vbNewLine
    End If
    
    Dim Client As New WebClient

    Client.BaseUrl = Worksheets(label).Range("url").Value
    Client.TimeoutMs = 75000
    
    Dim Request As New WebRequest
    Request.RequestFormat = WebFormat.json

    Request.Resource = "data/{label}"
    
    Request.Method = WebMethod.HttpPost
    Dim bndry As String
    bndry = "----WebKitFormBoundaryW34T6HD7JCW8"
    Request.ContentType = "multipart/form-data; boundary=" & bndry
        
    Request.AddHeader "x-token", Worksheets(label).Range("xtoken").Value
    
    Request.AddUrlSegment "label", label
    Request.AddQuerystringParam "runNano", "false"
    Request.AddQuerystringParam "fileType", "csv"
    Request.AddQuerystringParam "gzip", "false"
    Request.AddQuerystringParam "results", ""
    Request.AddQuerystringParam "api-tenant", Worksheets(label).Range("apitenant").Value
    
    Dim PostBody As String, Response As WebResponse, json As Object, dataSubsection As Long, i As Long, j As Long
    
    ''''''''''''''''
    dataSubsection = 1
    
    Request.AddQuerystringParam "appendData", "false"
    ' Do While dataSubsection <= row
    
    arrString = ""
    
    ' MsgBox dataSubsection
    For i = dataSubsection To row ' WorksheetFunction.Min(row, 30000 / col)
        tmpStr = ""
        For j = 1 To col
            tmpStr = tmpStr & "," & CStr(Selection.Cells(i, j))
        Next j
        tmpStr = Right(tmpStr, Len(tmpStr) - 1)
        arrString = arrString & tmpStr
        If i = row Then ' Or 30000 / col
            arrString = arrString & returnStr
        Else
            arrString = arrString & ","
        End If
    Next i
    PostBody = "--" & bndry & returnStr _
    & "Content-Disposition: form-data; name=""data""; filename=""example.csv""" & returnStr _
    & "Content-Type: application/vnd.ms-excel" & returnStr & returnStr _
    & arrString & returnStr _
    & "--" & bndry & "--" & returnStr

    Request.Body = PostBody
    
    Request.ResponseFormat = WebFormat.json
    Set Response = Client.Execute(Request)
    
    If Response.StatusCode <> 200 Then
        On Error GoTo JSONErr
        Set json = JsonConverter.ParseJson(Response.Content)
    
        MsgBox "NANO ERROR:" & vbNewLine & "   " & json("message")
        LoadData = False
        Exit Function
    End If
    ' dataSubsection = dataSubsection + 30000 / col
    
    ' Request.AddQuerystringParam "appendData", "true"
    ' Loop
    
    ''''''''''''''''
    
    Range("status").Value = "finished"
    
Exit Function
    
Err:
    Select Case Err.Number
        Case 6
            MsgBox "Load data failed: data size too big"
            LoadData = False
            Exit Function
            
        Case Else
            MsgBox "Load data failed: " & Err.Description
            LoadData = False
            Exit Function
    End Select

JSONErr:
    MsgBox "Response error: load data"
    LoadData = False
    Exit Function

End Function


Private Function RunNano() As Boolean
    RunNano = True
    On Error GoTo Err
    If Not (LoadData) Then
        RunNano = False
        Exit Function
    End If
    
    Dim label As String
    label = Range("currentNano").Value
    
    Range("status").Value = "running nano"
    Dim Client As New WebClient
    
    Client.BaseUrl = Worksheets(label).Range("url").Value
    Client.TimeoutMs = 75000
    
    Dim Request As New WebRequest
    Request.Resource = "nanoRun/{label}"
    Request.Method = WebMethod.HttpPost
    Request.AddUrlSegment "label", label
    Request.AddQuerystringParam "api-tenant", Worksheets(label).Range("apitenant").Value
    Request.AddHeader "x-token", Worksheets(label).Range("xtoken").Value

    Dim Response As WebResponse
    Set Response = Client.Execute(Request)
    
    Dim json As Object
    If Response.StatusCode <> 200 Then
        On Error GoTo JSONErr
        Set json = JsonConverter.ParseJson(Response.Content)
        
        MsgBox "NANO ERROR:" & vbNewLine & "   " & json("message")
        RunNano = False
    Else
        If GetStatus Then
                ' On Error GoTo Err
                ExportAnomalies
            MsgBox "Clustering successful"
        Else
            RunNano = False
            Exit Function
        End If
    End If
    
    Range("status").Value = "finished"
Exit Function
    
Err:
    MsgBox "Clustering failed: " & Err.Description
    RunNano = False
    Exit Function

JSONErr:
    MsgBox "Response error: clustering"
    RunNano = False
    Exit Function
    
End Function

Private Function ExportAnomalies() As Boolean
    Dim results As Variant, label As String, t As Integer, startRow As Integer
    
'    label = "Anomalies"
'    If WorksheetExists(label) Then
'        Worksheets(label).Cells.Clear
'    Else
'        Set NewSheet = Worksheets.Add(After:=Worksheets("BoonNano"))
'        NewSheet.Name = label
'    End If
'    Worksheets("BoonNano").Activate
'
     Set results = GetResults
'
'    numAnomalies = 0
'    For i = 1 To results("RI").Count
'        If results("RI")(i) >= Worksheets("BoonNano").Range("anomalyIndex").Value Then
'            numAnomalies = numAnomalies + 1
'            For j = 1 To Worksheets("BoonNano").Range("streamingWindowSize").Value
'
'
'            Next j
'            Selection.Rows(i).Copy
'            Worksheets(label).Range("$A$" & numAnomalies).PasteSpecial (xlPasteValues)
'        End If
'    Next i
'    Worksheets("BoonNano").Range("numAnomalies").Value = numAnomalies
    
    label = "Results"
    If WorksheetExists(label) Then
        ' Worksheets(label).Cells.Clear
        startRow = Worksheets(label).Cells(Rows.Count, 1).End(xlUp) + 1
        ' Worksheets(label).Cells(Rows.Count, 1).End (xlUp)
    Else
        Set NewSheet = Worksheets.Add(After:=Worksheets("BoonNano"))
        NewSheet.Name = label
        Worksheets("Results").Columns("G").Select
        ActiveWindow.FreezePanes = True
        startRow = 1
        Worksheets(label).Rows(1).Font.Bold = True
        Worksheets(label).Cells(1, 1) = "Pattern Number"
        Worksheets(label).Cells(1, 2) = "Cluster ID"
        Worksheets(label).Cells(1, 3) = "Anomaly Index"
        Worksheets(label).Cells(1, 4) = "Smoothed Anomaly Index"
        Worksheets(label).Cells(1, 5) = "Frequency Index"
        Worksheets(label).Cells(1, 6) = "Distance Index"
    End If
    

    For i = 1 To results("RI").Count
        Worksheets(label).Cells(i + startRow, 1) = i + startRow - 1
        Worksheets(label).Cells(i + startRow, 2) = results("ID")(i)
        Worksheets(label).Cells(i + startRow, 3) = results("RI")(i)
        Worksheets(label).Cells(i + startRow, 4) = results("SI")(i)
        Worksheets(label).Cells(i + startRow, 5) = results("FI")(i)
        Worksheets(label).Cells(i + startRow, 6) = results("DI")(i)
    Next i
    With Worksheets(label).Columns("A:F")
        .AutoFit
        .HorizontalAlignment = xlCenter
    End With
    
    Worksheets("BoonNano").Activate

End Function

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

