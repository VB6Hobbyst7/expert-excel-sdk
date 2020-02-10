Attribute VB_Name = "PageSetup"
Sub Boonnano()
    If Not (BoonHeaders) Then
        Exit Sub
    End If
    If Not (ParamHeaders) Then
        Exit Sub
    End If
    OpenButton
End Sub

Private Sub OpenButton()
    Dim btn As Button
    Set t = Worksheets("BoonNano").Range("C2:C3")
    Set btn = Worksheets("BoonNano").Buttons.Add(t.Left + 15, t.Top, t.Width - 15, t.Height)
    With btn
        .Caption = "Open"
        .Name = "openBtn"
        .OnAction = "GetButtons"
    End With

End Sub

Private Sub GetButtons()
    If IsEmpty(Worksheets("BoonNano").Range("user")) = False And IsEmpty(Worksheets("BoonNano").Range("currentNano")) = False Then
        Dim list As String
        list = Application.Run("management.GetUsers")
        If list <> "False" Then
        Worksheets("BoonNano").Range("user").Validation.Modify _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Formula1:=Application.Run("management.GetUsers")
        Else
            GoTo Err
        End If
        
        On Error GoTo Err
        AutotuneButton
        ByFeatureCheckbox
        ConfigureButton
        ' ResetBufferButton
        ' ResultsButton
        Worksheets("BoonNano").Shapes("openBtn").Delete
        CloseButton
        On Error GoTo Err
        Application.Run ("management.OpenNano")
    Else
        MsgBox "Enter the user and nano label"
    End If
    Exit Sub
    
Err:
    MsgBox "Cannot create buttons"
    CloseCleanup
    Exit Sub
    
End Sub


Private Sub ByFeatureCheckbox()
    Dim chbx As CheckBox
    Set t = Worksheets("BoonNano").Range("E3:F3")
    Set chbx = Worksheets("BoonNano").CheckBoxes.Add(Left:=t.Left + 10, Top:=t.Top - 3, Width:=t.Width - 50, Height:=6)
    With chbx
        .Name = "ByFeature"
        .Caption = "By Feature"
    End With
End Sub

Private Sub CloseButton()
    Dim btn As Button
    Set t = Worksheets("BoonNano").Range("C2:C3")
    Set btn = Worksheets("BoonNano").Buttons.Add(t.Left + 15, t.Top, t.Width - 15, t.Height)
    With btn
        .Caption = "Close"
        .Name = "closeBtn"
        .OnAction = "CloseCleanup"
    End With
End Sub

Private Sub CloseCleanup()
    On Error Resume Next
    Application.Run ("management.CloseNano")
    
    Worksheets("BoonNano").Range("currentNano").ClearContents
    With Worksheets("BoonNano")
        .Shapes("Autotune").Delete
        .Shapes("ByFeature").Delete
        .Shapes("Cluster").Delete
        ' .Shapes("Results").Delete
        .Shapes("Configure").Delete
        ' .Shapes("Reset").Delete
        .Shapes("closeBtn").Delete
        .Range("status") = ""
        .Range("byteProcess,byteBuffer,byteWritten,numClusters,totalInferences,avgClusterTime,numAnomalies").Value = 0
    End With
    OpenButton

End Sub

Private Sub AutotuneButton()
    Dim btn As Button
    Set t = Worksheets("BoonNano").Range("E2:F2")
    Set btn = Worksheets("BoonNano").Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
        .Name = "Autotune"
        .Caption = "Autotune Selection"
        .OnAction = "config.AutotuneConfig"
    End With
End Sub

Private Sub ClusterButton()
    Dim btn As Button
    Set t = Worksheets("BoonNano").Range("H2:I3")
    Set btn = Worksheets("BoonNano").Buttons.Add(t.Left, t.Top - Range("H2").Height, t.Width, t.Height + Range("H2").Height)
    With btn
        .Name = "Cluster"
        .Caption = "Cluster Selection"
        .OnAction = "results.RunNano"
        .Enabled = False
    End With
End Sub

Private Sub ConfigureButton()
    Dim btn As Button
    Set t = Worksheets("BoonNano").Range("E2:F2")
    Set btn = Worksheets("BoonNano").Buttons.Add(t.Left, t.Top - t.Height, t.Width, t.Height)
    With btn
        .Name = "Configure"
        .Caption = "Configure"
        .OnAction = "config.SetConfig"
    End With
End Sub

Private Sub ResetBuffer()
    On Error GoTo Err
    Application.Run ("management.CloseNano")
    Application.Run ("management.OpenNano")
    Application.Run ("results.GetBufferStatus")
    Range("numClusters").Value = 0
    Exit Sub

Err:
    Exit Sub
    
End Sub

Private Sub ResetBufferButton()
    Dim btn As Button
    Set t = Worksheets("BoonNano").Range("H3:I3")
    Set btn = Worksheets("BoonNano").Buttons.Add(t.Left + 10, t.Top + 2, t.Width - 20, t.Height - 4)
    With btn
        .Name = "Reset"
        .Caption = "Reset Buffer"
        .OnAction = "ResetBuffer"
    End With
End Sub


Private Sub ResultsButton()
    Dim btn As Button
    Set t = Worksheets("BoonNano").Range("K2:L2")
    Set btn = Worksheets("BoonNano").Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
        .Name = "Results"
        .Caption = "Export Anomalies"
        .OnAction = "results.ExportAnomalies"
    End With
End Sub

Private Function ParamHeaders() As Boolean
    ParamHeaders = True
    On Error GoTo Err
    Rows("4:7").Insert Shift:=xlShiftDown
    With Rows("4:7")
        .RowHeight = 17
        .Interior.Color = RGB(235, 235, 235)
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End With
    Range("B4:B7,A4,A8:A18,A20:A24").Font.Bold = True
        
    With Worksheets("BoonNano")
    
        ' CLUSTER STATUS
        With .Range("A4")
            .HorizontalAlignment = xlCenter
            .Value = "Cluster status"
            .Font.Size = 14
        End With
        
        .Range("A5:A7").Merge
        With .Range("A4:A7").Borders
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
        .Range("A4").Borders(xlEdgeBottom).Weight = xlThin
            
        With .Range("A5")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Name = "status"
            .Value = "finished"
            .Font.Size = 14
            .FormatConditions.Add Type:=xlExpression, Formula1:="=And(A5<>""finished"", ISBLANK(A5)=FALSE)"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1)
                With .Interior
                    .PatternColorIndex = xlAutomatic
                    .Color = RGB(255, 0, 0)
                End With
            .StopIfTrue = False
            End With
        End With
        
        ' CONFIG PARAMETERS
        .Range("B4").Value = "Weight"
        .Range("B5").Value = "Max"
        .Range("B6").Value = "Min"
        .Range("B7").Value = "Label"
        With .Rows("7").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
    
        .Range("A8").Value = "Percent Variation"
        .Range("A9").Value = "Numeric Type"
        .Range("A10").Value = "Streaming Window"
        .Range("A11").Value = "Accuracy"
        .Range("A12").Value = "Feature Count"
        .Range("A13").Value = "Anomaly Threshold"
    
        .Range("B8").Name = "percentVariation"
        .Range("B9").Name = "numericFormat"
        .Range("B10").Name = "streamingWindowSize"
        .Range("B11").Name = "accuracy"
        .Range("B12").Name = "numFeatures"
        .Range("B13").Name = "anomalyIndex"
    
        ' BUFFER DATA
        .Range("A15:B15").Merge
        With .Range("A15")
            .Value = "Data Buffer"
            .Font.Size = 14
            .HorizontalAlignment = xlCenter
        End With
        .Range("A16").Value = "Bytes in buffer"
        .Range("A17").Value = "Bytes processed"
        .Range("A18").Value = "Bytes written"
        
        .Range("B16").Name = "byteBuffer"
        .Range("B17").Name = "byteProcess"
        .Range("B18").Name = "byteWritten"
        
        ' CLUSTER SUMMARY
        .Range("A20:B20").Merge
        With .Range("A20")
            .Value = "Cluster Summary"
            .Font.Size = 14
            .HorizontalAlignment = xlCenter
        End With
        .Range("A21").Value = "Number of clusters"
        .Range("B21").Name = "numClusters"
        
        .Range("A22").Value = "Clustered inferences"
        .Range("B22").Name = "totalInferences"
        
        .Range("A23").Value = "Average cluster time (" & ChrW(181) & "s)"
        .Range("B23").Name = "avgClusterTime"
        
        .Range("A24").Value = "Number of Anomalies"
        .Range("B24").Name = "numAnomalies"

        ' color param headers
        With .Range("A8:B13,A15:B18,A20:B24")
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .Interior.Color = RGB(235, 235, 235)
        End With
        
    End With
    Exit Function
    
Err:
    MsgBox "Error with adding rows/columns"
    ParamHeaders = False
    Exit Function
End Function


Private Function BoonHeaders() As Boolean
    BoonHeaders = True
    On Error GoTo Err
    ActiveWorkbook.ActiveSheet.Name = "BoonNano"
    Worksheets("BoonNano").Rows("1:3").Insert Shift:=xlShiftDown, CopyOrigin:=xlInsertFormatOriginConstant
    Worksheets("BoonNano").Columns("A:B").Insert Shift:=xlShiftRight, CopyOrigin:=xlInsertFormatOriginConstant
    Worksheets("BoonNano").Range("A1:A3").Font.Bold = True
    
    ' HEADER
    ' applies to all header
    Worksheets("BoonNano").Rows("1:3").Interior.Color = RGB(197, 217, 241)
    With Worksheets("BoonNano").Rows("3").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' add text formating
    With Worksheets("BoonNano").Range("A1")
        ' .Value = "BoonLogic"
        .HorizontalAlignment = xlCenter
    End With
    Worksheets("BoonNano").Rows("1").RowHeight = 54
    Worksheets("BoonNano").Columns("A").ColumnWidth = 22
    Worksheets("BoonNano").Range("A1").Font.Size = 28
    With Worksheets("BoonNano").Range("A2")
        .Value = "User"
        .HorizontalAlignment = xlRight
        .Font.Size = 16
    End With
    
    Dim pic As String, t As Range
    Set t = Range("A1")
    pic = "https://raw.githubusercontent.com/boonlogic/boonlogic-rest-api/master/images/BoonLogic.png"
    On Error GoTo NoLogo
    Set Boonlogo = Worksheets("BoonNano").Pictures.Insert(pic)
    With Boonlogo
        .ShapeRange.LockAspectRatio = msoTrue
        .Width = t.Width
        .Height = t.Height
        .Top = t.Top
        .Left = t.Left
    End With
    
Headers:
    
    Worksheets("BoonNano").Columns("B").ColumnWidth = 11.5
    With Worksheets("BoonNano").Range("B2")
        .HorizontalAlignment = xlCenter
        .Name = "user"
        .Value = "default"
        Dim list As String
        list = Application.Run("management.GetUsers")
        If list <> "False" Then
            .Validation.Add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Formula1:=list
        Else
            .Validation.Add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Formula1:="default"
        End If
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        .FormatConditions.Add Type:=xlExpression, Formula1:="=ISBLANK(B2)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1)
            With .Interior
                .PatternColorIndex = xlAutomatic
                .Color = RGB(255, 0, 0)
            End With
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        .StopIfTrue = False
        End With
    End With
    
            
    With Worksheets("BoonNano").Range("A3")
        .Value = "Nano label"
        .Font.Size = 16
        .HorizontalAlignment = xlRight
    End With
    With Worksheets("BoonNano").Range("B3")
        .HorizontalAlignment = xlCenter
        .Name = "currentNano"
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        .FormatConditions.Add Type:=xlExpression, Formula1:="=ISBLANK(B3)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1)
            With .Interior
                .PatternColorIndex = xlAutomatic
                .Color = RGB(255, 0, 0)
            End With
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        .StopIfTrue = False
        End With
    End With
    Exit Function
    
Err:
    MsgBox "Error with adding rows/columns " & Err.Description
    BoonHeaders = False
    Exit Function
    
NoLogo:
    On Error GoTo -1
    With Worksheets("BoonNano").Range("A1")
        .Value = "BoonLogic"
        .VerticalAlignment = xlCenter
    End With
    On Error GoTo 0
    GoTo Headers
    
End Function
