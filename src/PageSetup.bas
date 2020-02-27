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
    Worksheets("BoonNano").Range("B1").Value = "User"
    Worksheets("BoonNano").Range("B2").Value = "Nano label"
    Set t = Worksheets("BoonNano").Range("D1:D2")
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
    Set t = Worksheets("BoonNano").Range("A6")
    Set chbx = Worksheets("BoonNano").CheckBoxes.Add(Left:=t.Left + 10, Top:=t.Top - 3, Width:=t.Width - 50, Height:=6)
    With chbx
        .Name = "ByFeature"
        .Caption = "By Feature"
    End With
End Sub

Private Sub CloseButton()
    Dim btn As Button
    Worksheets("BoonNano").Range("B1:B2").Value = ""
    Set t = Worksheets("BoonNano").Range("C1:C2")
    Set btn = Worksheets("BoonNano").Buttons.Add(t.Left, t.Top, t.Width, t.Height)
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
        .Shapes("Reset").Delete
        .Shapes("closeBtn").Delete
        .Range("status") = ""
        .Range("byteProcess,byteBuffer,byteWritten,numClusters,totalInferences,avgClusterTime").Value = 0
    End With
    OpenButton

End Sub

Private Sub AutotuneButton()
    Dim btn As Button
    Set t = Worksheets("BoonNano").Range("A5")
    Set btn = Worksheets("BoonNano").Buttons.Add(t.Left + 20, t.Top + 2, t.Width - 40, t.Height - 4)
    With btn
        .Name = "Autotune"
        .Caption = "Autotune"
        .OnAction = "config.AutotuneConfig"
    End With
End Sub

Private Sub ClusterButton()
    Dim btn As Button
    Set t = Worksheets("BoonNano").Range("E1:E2")
    Set btn = Worksheets("BoonNano").Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
        .Name = "Cluster"
        .Caption = "Cluster Now"
        .OnAction = "results.RunNano"
        .Enabled = False
    End With
End Sub

Private Sub ConfigureButton()
    Dim btn As Button
    Set t = Worksheets("BoonNano").Range("A4")
    Set btn = Worksheets("BoonNano").Buttons.Add(t.Left + 20, t.Top + 2, t.Width - 40, t.Height - 4)
    With btn
        .Name = "Configure"
        .Caption = "Manual"
        .OnAction = "config.SetConfig"
    End With
End Sub

Private Sub ResetBuffer()
    On Error GoTo Err
    Application.Run ("management.CloseNano")
    Application.Run ("management.OpenNano")
    ' Application.Run ("results.GetBufferStatus")
    On Error Resume Next
    Worksheets("BoonNano").Shapes("Cluster").Delete
    Range("C3:XFD6,percentVariation,numericFormat,streamingWindowSize,accuracy,numFeatures,anomalyIndex,numClusters,totalInferences,avgClusterTime") = ""
    Exit Sub

Err:
    MsgBox "ERROR: " & Err.Description
    Exit Sub
    
End Sub

Private Sub ResetBufferButton()
    Dim btn As Button
    Set t = Worksheets("BoonNano").Range("F2")
    Set btn = Worksheets("BoonNano").Buttons.Add(t.Left + 2, t.Top + 2, t.Width - 4, t.Height - 4)
    With btn
        .Name = "Reset"
        .Caption = "Reset"
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
    Rows("3").Delete
    With Rows("3:6")
        .RowHeight = 17
        .Interior.Color = RGB(198, 224, 180)
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End With
        
    With Worksheets("BoonNano")
        .Range("A3,B3:B6,A7:A17").Font.Bold = True
        With .Range("A3")
            .HorizontalAlignment = xlCenter
            .Value = "Configure Parameters"
        End With
    

        ' CONFIG PARAMETERS
        .Range("B3").Value = "Weight"
        .Range("B4").Value = "Max"
        .Range("B5").Value = "Min"
        .Range("B6").Value = "Label"
        With .Rows("6").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
        .Range("B3:B6").HorizontalAlignment = xlCenter
    
        .Range("A7").Value = "Percent Variation"
        .Range("A8").Value = "Numeric Type"
        .Range("A9").Value = "Streaming Window"
        .Range("A10").Value = "Accuracy"
        .Range("A11").Value = "Feature Count"
        .Range("A12").Value = "Anomaly Threshold"
    
        .Range("B7").Name = "percentVariation"
        .Range("B8").Name = "numericFormat"
        .Range("B9").Name = "streamingWindowSize"
        .Range("B10").Name = "accuracy"
        .Range("B11").Name = "numFeatures"
        .Range("B12").Name = "anomalyIndex"
    
'        ' BUFFER DATA
'        .Range("A15:B15").Merge
'        With .Range("A15")
'            .Value = "Data Buffer"
'            .Font.Size = 14
'            .HorizontalAlignment = xlCenter
'        End With
'        .Range("A16").Value = "Bytes in buffer"
'        .Range("A17").Value = "Bytes processed"
'        .Range("A18").Value = "Bytes written"
'
'        .Range("B16").Name = "byteBuffer"
'        .Range("B17").Name = "byteProcess"
'        .Range("B18").Name = "byteWritten"
'
        ' CLUSTER SUMMARY
        .Range("A14:B14").Merge
        With .Range("A14")
            .Value = "Nano Status"
            .Font.Size = 14
            .HorizontalAlignment = xlCenter
        End With
        .Range("A15").Value = "Number of clusters"
        .Range("B15").Name = "numClusters"
    
        .Range("A16").Value = "Patterns processed"
        .Range("B16").Name = "totalInferences"
        
        .Range("A17").Value = "Average cluster time (" & ChrW(181) & "s)"
        .Range("B17").Name = "avgClusterTime"
        
'        .Range("A18").Value = "Number of Anomalies"
'        .Range("B18").Name = "numAnomalies"

        ' color param headers
        With .Range("A7:B12")
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .Interior.Color = RGB(198, 224, 180)
        End With
        
        ' nano status headers
        With .Range("A14:B17")
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .Interior.Color = RGB(197, 217, 241)
        End With
        Range("C7").Select
        ActiveWindow.FreezePanes = True
        
    End With
    Exit Function
    
Err:
    MsgBox "Error with adding rows/columns"
    ParamHeaders = False
    Exit Function
End Function


Private Function BoonHeaders() As Boolean

    Application.ReferenceStyle = xlA1
    
    BoonHeaders = True
    On Error GoTo Err
    ActiveWorkbook.ActiveSheet.Name = "BoonNano"
    Worksheets("BoonNano").Rows("1:3").Insert Shift:=xlShiftDown, CopyOrigin:=xlInsertFormatOriginConstant
    Worksheets("BoonNano").Columns("A:B").Insert Shift:=xlShiftRight, CopyOrigin:=xlInsertFormatOriginConstant
    Worksheets("BoonNano").Range("A1:A2").Font.Bold = True
    
    ' HEADER
    ' applies to all header
    With Worksheets("BoonNano").Rows("1:2")
        .Interior.Color = RGB(197, 217, 241)
        .RowHeight = 19
    End With
    With Worksheets("BoonNano").Rows("2").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' add text formating
    With Worksheets("BoonNano").Range("A1:A2")
        ' .Value = "BoonLogic"
        .Merge
        .HorizontalAlignment = xlCenter
    End With
    ' Worksheets("BoonNano").Rows("1").RowHeight = 54
    Worksheets("BoonNano").Columns("A").ColumnWidth = 22
    Worksheets("BoonNano").Range("A1").Font.Size = 28
    
    Dim pic As String, t As Range
    Set t = Range("A1:A2")
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

    Worksheets("BoonNano").Columns("B").ColumnWidth = 13

    With Worksheets("BoonNano").Range("B1")
        ' .Value = "User"
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .Font.Size = 16
    End With
    
    With Worksheets("BoonNano").Range("C1")
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
        .FormatConditions.Add Type:=xlExpression, Formula1:="=ISBLANK(C1)"
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
    
            
    With Worksheets("BoonNano").Range("B2")
        ' .Value = "Nano label"
        .Font.Size = 16
        .HorizontalAlignment = xlRight
    End With
    With Worksheets("BoonNano").Range("C2")
        .HorizontalAlignment = xlCenter
        .Name = "currentNano"
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        .FormatConditions.Add Type:=xlExpression, Formula1:="=ISBLANK(C2)"
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
    
        ' CLUSTER STATUS
        With Worksheets("BoonNano").Range("H1:J1")
            .Merge
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .Value = "Cluster status"
            .Font.Size = 14
            .Borders(xlEdgeBottom).Weight = xlThin
        End With

        With Worksheets("BoonNano").Range("H2:J2")
           .Merge
           With .Borders
                .LineStyle = xlContinuous
               .Weight = xlThick
           End With

            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Name = "status"
            .Value = "finished"
            .Font.Size = 14
            .FormatConditions.Add Type:=xlExpression, Formula1:="=And(H2<>""finished"", ISBLANK(H2)=FALSE)"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1)
                With .Interior
                    .PatternColorIndex = xlAutomatic
                    .Color = RGB(255, 0, 0)
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
