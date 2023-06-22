Attribute VB_Name = "Module1"
Sub fill()
Attribute fill.VB_ProcData.VB_Invoke_Func = "d\n14"
           
      Selection.Value = Rnd() * 100
    
End Sub
Sub Generate_butt()
    Dim Num As Variant
    ActiveSheet.Cells.ClearContents
    Num = InputBox("How many cities do you want to visite ? ", "Insert Your Answer ")
    Cells(1, 1) = "The Number Of Cities"
    Cells(2, 1) = Num
    Cells(1, 2) = "The Cities"
    Cells(1, 3) = "X"
    Cells(1, 4) = "Y"
    
    
    Call Refresh
    Call Table
    
End Sub
Sub Update()
    Dim varNum As Integer
    'copy process
    Cells(100, 100) = "=Count(" & Range(Range("C2"), Range("C2").End(xlDown)).Address & ")"
    varNum = Cells(100, 100)
    Dim cpy As Variant
    ReDim cpy(1 To Cells(100, 100) * 2)
    
    Dim k As Integer

    
    
    For k = 1 To varNum * 2
        ' copy X coordinates
        cpy(k) = Range("C2:D" & varNum).Cells(k)
        Debug.Print cpy(k)
    Next k
    

    ActiveSheet.Cells.ClearContents
    
    Dim MyChart As Object    ' Create object variable.
      Set MyChart = ActiveSheet.ChartObjects  ' Create valid object reference.
      MyChartCount = MyChart.Count
    
    If MyChartCount > 0 Then
        MyChart.Delete
    End If
    
    For k = 1 To varNum * 2
        ' copy X coordinates
        Range("C2:D" & varNum).Cells(k) = cpy(k)
        
        
    Next k
    
    Cells(2, 1) = varNum
    Debug.Print Cells(2, 1)
    Cells(1, 1) = "The Number Of Cities"
    Cells(1, 2) = "The Cities"
    Cells(1, 3) = "X"
    Cells(1, 4) = "Y"
    
    For i = 2 To Cells(2, 1) + 1
        Cells(i, 2) = "City " & i - 1
    Next i
    
    Call tt(2, Cells(2, 1), 1, 2)
    Call Table
    
End Sub
Sub Refresh()
    
   
    
    Num = Cells(2, 1)
    For i = 2 To Num + 1
        Cells(i, 2) = "City " & i - 1
        Cells(i, 3) = Rnd() * 100
        Cells(i, 4) = Rnd() * 100
    Next i
    
    
    
    Call tt(2, Cells(2, 1), 1, 2)
   
End Sub
Sub Chart()
      Dim MyChart As Object    ' Create object variable.
      Set MyChart = ActiveSheet.ChartObjects  ' Create valid object reference.
      MyChartCount = MyChart.Count
    
    If MyChartCount > 0 Then
        MyChart.Delete
    End If
    
    Range("C2:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select

    ActiveChart.FullSeriesCollection(1).MarkerSize = 10
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MaximumScale = 100
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MajorUnit = 10
    
    
    ActiveChart.Parent.Height = 423.8
    ActiveChart.Parent.Width = 431

    
    Call addseries
    
    Dim s As String
    s = "={"
    For i = 1 To Cells(2, 1)
    
        s = s & i & ","
    
    Next i
    s = Left(s, Len(s) - 1) & "}"
    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.SetElement (msoElementDataLabelTop)
    ActiveChart.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange. _
        InsertChartField msoChartFieldRange, s, 0

    ActiveChart.FullSeriesCollection(1).DataLabels.ShowRange = True
    ActiveChart.FullSeriesCollection(1).DataLabels.ShowValue = False
        With ActiveChart.FullSeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "Inter ExtraBold"
        .NameFarEast = "Inter ExtraBold"
        .Name = "Inter ExtraBold"
    End With
    
End Sub
Sub addseries()

    Call tt(2, Cells(2, 1), 1, 2)
    
End Sub

Sub tt(i As Integer, j As Integer, n As Integer, Sn As Integer)

    If i <= j Then
        For cnt = i To j
        ' [do the thing whatever]
           If ActiveSheet.ChartObjects.Count > 0 Then
        
                ActiveChart.SeriesCollection.NewSeries
                ActiveChart.FullSeriesCollection(Sn).XValues = "=Sheet1!" + Range(Range("C2"), Range("C2").End(xlDown)).Cells(n).Address + "," + "Sheet1!" + Range(Range("C2"), Range("C2").End(xlDown)).Cells(cnt).Address
                ActiveChart.FullSeriesCollection(Sn).Values = "=Sheet1!" + Range(Range("D2"), Range("D2").End(xlDown)).Cells(n).Address + "," + "Sheet1!" + Range(Range("D2"), Range("D2").End(xlDown)).Cells(cnt).Address
       
                ActiveChart.FullSeriesCollection(Sn).Select
                With Selection.Format.Line
                    .Visible = msoTrue
                    .ForeColor.ObjectThemeColor = msoThemeColorText2
                    .ForeColor.TintAndShade = 0
                    .ForeColor.Brightness = 0.8000000119
                End With
        
          
            Else
        
        ' Constraint coefficients
        
                Xi = Range(Range("C2"), Range("C2").End(xlDown)).Cells(n).Value
                Xf = Range(Range("C2"), Range("C2").End(xlDown)).Cells(cnt).Value
        
                Yi = Range(Range("D2"), Range("D2").End(xlDown)).Cells(n).Value
                Yf = Range(Range("D2"), Range("D2").End(xlDown)).Cells(cnt).Value
        

                Range("F2").Offset(cnt - 1, n - 1) = Sqr((Xf - Xi) ^ (2) + (Yf - Yi) ^ (2))
                Range("F2").Offset(n - 1, cnt - 1) = Range("F2").Offset(cnt - 1, n - 1)
        
                End If
            
            Sn = Sn + 1
        
        
        Next cnt
        Call tt(n + 2, Cells(2, 1), n + 1, Sn)
    End If
End Sub
    
Sub Table()
    Dim rc As String
    Dim urc As String
    For i = 0 To Cells(2, 1) - 1
        For j = 0 To Cells(2, 1) - 1
            
            If Not ((i = 0 Or j = 0) Or j = i) Then
                ' tap here the eleminating constraint
                
                Cells(2, 5 + 2 * Cells(2, 1) + 5).Offset(i, j).FormulaR1C1 = "=R[" & -1 & "]C[" & -j - 1 & "]" & "-" & "R[" & j - i - 1 & "]C[" & -j - 1 & "]" & "+" & 1 & "+ (" & Cells(2, 1) & "*" & "(" & "R[" & 0 & "]C[" & -3 - Cells(2, 1) & "]" & "-" & 1 & ")" & ")"
                'Debug.Print i & "  " & j
                
            End If
            
            
            
        Next j
    Next i
    
   
    
    
    For i = 1 To Cells(2, 1)
        
        Cells(1, 5 + i) = "City " & i
        Cells(i + 1, 5) = "City " & i
        
        Cells(1, 5 + Cells(2, 1) + 1 + i) = "City " & i
        Cells(i + 1, 5 + Cells(2, 1) + 1) = "City " & i
        
        ' sub tour vars
        If i <> 1 Then
            Cells(i, 5 + 2 * Cells(2, 1) + 3) = "U" & i
        End If
        
        
        
        ' the diagonal constraint :
        rc = "R[" & -i & "]C[" & -i & "]" & "+" & rc
        
    Next i
    
    Cells(1, 5 + 2 * Cells(2, 1) + 4) = "ST Variables"
    
    
    With Cells(1 + Cells(2, 1) + 1, 5 + Cells(2, 1) + 1)
        .Value = "Constraints"
        .Offset(0, Cells(2, 1) + 1).FormulaR1C1 = "=sum(" & Left(rc, Len(rc) - 1) & ")"
        With .Offset(-1 * (Cells(2, 1) + 1), Cells(2, 1) + 1)
            .Value = "Constraints"
            .Offset(1, 0).FormulaR1C1 = "=SUM(RC[" & (-1) * Cells(2, 1) & "]:RC[-1])"
            .Offset(1, 0).AutoFill Destination:=Range(Cells(1 + Cells(2, 1) + 1, 5 + Cells(2, 1) + 1).Offset(-1 * (Cells(2, 1) + 1), Cells(2, 1) + 1).Offset(1, 0).Address & ":" & Cells(1 + Cells(2, 1) + 1, 5 + Cells(2, 1) + 1).Offset(-1 * (Cells(2, 1) + 1), Cells(2, 1) + 1).Offset(Cells(2, 1), 0).Address), Type:=xlFillDefault
        End With
        .Offset(0, 1).FormulaR1C1 = "=sum(R[" & (-1) * Cells(2, 1) & "]C:R[-1]C)"
        .Offset(0, 1).AutoFill Destination:=Range(Cells(1 + Cells(2, 1) + 1, 5 + Cells(2, 1) + 1).Offset(0, 1).Address & ":" & Cells(1 + Cells(2, 1) + 1, 5 + Cells(2, 1) + 1).Offset(0, Cells(2, 1)).Address), Type:=xlFillDefault
    End With
    
    ' the objective function
    
    Range("A3").Value = "Objective Function"
    Range("A4").FormulaR1C1 = "=SUMPRODUCT(" & "R[-2]C[5]:" & "R[" & (Cells(2, 1) - 3) & "]C[" & (4 + Cells(2, 1)) & "]," & "R[-2]C[" & (5 + Cells(2, 1) + 1) & "]:" & "R[" & (Cells(2, 1) - 3) & "]C[" & (4 + 2 * Cells(2, 1) + 1) & "])"
    
End Sub

Sub Solve()
    SolverReset
    Dim rangerover As String
    Dim subtoursmover As String
    rangerover = Cells(2, 4 + Cells(2, 1) + 2 + 1).Address & ":" & Cells(2, 4 + Cells(2, 1) + 2 + 1).Offset(Cells(2, 1) - 1, Cells(2, 1) - 1).Address
    subtoursmover = Cells(2, 5 + 2 * Cells(2, 1) + 4).Address & ":" & Cells(2, 5 + 2 * Cells(2, 1) + 4).Offset(Cells(2, 1) - 2, 0).Address
    SolverOk SetCell:="$A$4", MaxMinVal:=2, ValueOf:=0, ByChange:=rangerover & "," & subtoursmover, _
        Engine:=1, EngineDesc:="Simplex LP"

    For i = 0 To Cells(2, 1) - 1
        For j = 0 To Cells(2, 1) - 1
            
            
            SolverAdd CellRef:=Cells(2, 4 + Cells(2, 1) + 2 + 1).Offset(i, j).Address, Relation:=5, FormulaText:="binary"
            
           If Not ((i = 0 Or j = 0) Or j = i) Then
                ' tap here the eleminating constraint
                
                SolverAdd CellRef:=Cells(2, 5 + 2 * Cells(2, 1) + 5).Offset(i, j).Address, Relation:=1, FormulaText:=0
                
                
           End If
            
        Next j
        
        SolverAdd CellRef:=Cells(2, 4 + Cells(2, 1) + 2 + 1).Offset(Cells(2, 1), i).Address, Relation:=2, FormulaText:="sum(1)"
        SolverAdd CellRef:=Cells(2, 4 + Cells(2, 1) + 2 + 1).Offset(i, Cells(2, 1)).Address, Relation:=2, FormulaText:="sum(1)"
        
        If i < Cells(2, 1) - 1 Then
            SolverAdd CellRef:=Cells(2, 5 + 2 * Cells(2, 1) + 4).Offset(i, 0).Address, Relation:=4, FormulaText:="integer"
            SolverAdd CellRef:=Cells(2, 5 + 2 * Cells(2, 1) + 4).Offset(i, 0).Address, Relation:=1, FormulaText:="sum(" & Cells(2, 1) - 1 & ")"
        End If
            
        
        
    Next i
    SolverAdd CellRef:=Cells(2, 4 + Cells(2, 1) + 2 + 1).Offset(Cells(2, 1), Cells(2, 1)).Address, Relation:=2, FormulaText:="0"
    
    'SolverSolve
    OpenSolver.RunOpenSolver , False
    Call graphsol
    
End Sub

Sub graphsol()
    Call Chart
    Sn = ActiveChart.SeriesCollection.Count + 1
    For i = 0 To Cells(2, 1) - 1
        For j = 0 To Cells(2, 1) - 1
            
            If Cells(2, 4 + Cells(2, 1) + 2 + 1).Offset(i, j) = 1 Then
                ActiveChart.SeriesCollection.NewSeries
                ActiveChart.FullSeriesCollection(Sn).XValues = "=Sheet1!" & Range(Range("C2"), Range("C2").End(xlDown)).Cells(i + 1).Address & "," & "Sheet1!" & Range(Range("C2"), Range("C2").End(xlDown)).Cells(j + 1).Address
                ActiveChart.FullSeriesCollection(Sn).Values = "=Sheet1!" & Range(Range("D2"), Range("D2").End(xlDown)).Cells(i + 1).Address & "," & "Sheet1!" & Range(Range("D2"), Range("D2").End(xlDown)).Cells(j + 1).Address
                
                
     ActiveChart.FullSeriesCollection(Sn).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(214, 220, 229)
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    Selection.Format.Line.EndArrowheadStyle = msoArrowheadTriangle
    With Selection.Format.Line
        .EndArrowheadLength = msoArrowheadLong
        .EndArrowheadWidth = msoArrowheadWide
    End With
                
                
                
                
                
                Sn = Sn + 1
            End If
            
        Next j
    Next i
End Sub

Sub NNH()
    Dim MinCost As Double
    Dim Conditions() As String
    ReDim Conditions(0 To Cells(2, 1) - 1)
    Dim k, n, i, j As Integer
    Dim NNHdeterminator As Double

    Dim NHHmin As Double
    NNHmin = 1E+99
    Dim NNHoptimal() As String
    ReDim NNHoptimal(0 To Cells(2, 1) - 1)
    
   For n = 0 To Cells(2, 1) - 1
    
        ReDim Conditions(0 To Cells(2, 1) - 1)
        MinCost = Cells(2, 4 + 2).Offset(n, (n + 1) Mod Cells(2, 1)).Value
        'Debug.Print "cost : " & MinCost
        
         
         k = 0
         Conditions(k) = "" & n & "," & (n + 1) Mod Cells(2, 1)
         For i = n To Cells(2, 1) - 1
             For j = 0 To Cells(2, 1) - 1
                 If i <> j Then
                     'Debug.Print TspCondition(Conditions, j)
                     
                     If (Cells(2, 4 + 2).Offset(i, j) < MinCost) And (TspCondition(Conditions, j)) Then
                         
                         Conditions(k) = "" & i & "," & j
                         
                         'Debug.Print "Condition : " & Conditions(k)
                         
                         MinCost = Cells(2, 4 + 2).Offset(i, j).Value
                     End If
                 End If
             Next j
             
            
             If Conditions(UBound(Conditions)) <> "" Then
            
                 Exit For
             End If
             
             'i = Right(Conditions(k), 1) - 1
             i = Mid(Conditions(k), InStr(Conditions(k), ",") + 1, Len(Conditions(k))) - 1
             
            ' Debug.Print "i : " & i
             Dim m As Integer
             Dim b As Boolean
             b = False
             For m = 0 To Cells(2, 1) - 1
                 
                 If m <> i + 1 And (TspCondition(Conditions, m)) Then
                     MinCost = Cells(2, 4 + 2).Offset(i + 1, m).Value
                     'Debug.Print "Lcost : " & MinCost
                     b = True
                     
                     Exit For
                     
                 End If
                 
             Next m
             
             
                If b = False Then
                    m = n
                    'Debug.Print "m is false"
                End If
             
            ' Debug.Print "mincost " & MinCost
             
             k = k + 1
             Conditions(k) = "" & i + 1 & "," & m
             'Debug.Print "bigCondition " & Conditions(k)
             
             
        Next i
         
         
         
            Debug.Print "*******************************"
            NNHdeterminator = 0
            For i = 0 To UBound(Conditions)
                NNHdeterminator = NNHdeterminator + Cells(2, 4 + 2).Offset(Mid(Conditions(i), 1, InStr(Conditions(i), ",") - 1), Mid(Conditions(i), InStr(Conditions(i), ",") + 1, Len(Conditions(i))))
                Debug.Print Mid(Conditions(i), 1, InStr(Conditions(i), ",") - 1) + 1 & "-->" & Mid(Conditions(i), InStr(Conditions(i), ",") + 1, Len(Conditions(i))) + 1
            Next i
             Debug.Print " The Nearest Neihber Cost : " & NNHdeterminator
     
         If NNHdeterminator < NNHmin Then
            NNHmin = NNHdeterminator
            For i = 0 To UBound(Conditions)
                NNHoptimal(i) = Conditions(i)
            Next i
         End If
         
    Next n
    
    
    Debug.Print "************************"
   Debug.Print "THE SHORTEST ROUTE IS :" & NNHmin
    
    Range("" & Cells(2, 4 + Cells(2, 1) + 2 + 1).Offset(0, 0).Address & ":" & Cells(2, 4 + Cells(2, 1) + 2 + 1).Offset(Cells(2, 1) - 1, Cells(2, 1) - 1).Address).Value = 0
       
    For i = 0 To Cells(2, 1) - 1
        Cells(2, 4 + Cells(2, 1) + 2 + 1).Offset(Mid(NNHoptimal(i), 1, InStr(NNHoptimal(i), ",") - 1), Mid(NNHoptimal(i), InStr(NNHoptimal(i), ",") + 1, Len(NNHoptimal(i)))) = 1
    
    Next i


     Call graphsol
    
End Sub

Function TspCondition(Condi() As String, j As Integer) As Boolean
    For i = 0 To UBound(Condi)
        If (Condi(i) <> "") Then
                If (j = Mid(Condi(i), InStr(Condi(i), ",") + 1, Len(Condi(i)))) Or (j = Mid(Condi(i), 1, InStr(Condi(i), ",") - 1)) Then
                    TspCondition = False
                    Exit Function ' Exit the function when condition is met
                End If
        End If
    Next i
    TspCondition = True ' Condition not met, return True
End Function


