VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptimiser 
   Caption         =   "SEMANTIC SYNDICATE"
   ClientHeight    =   10356
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   15552
   OleObjectBlob   =   "frmOptimiser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOptimiser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClear_Click()
    'CLEAR BUTTON
    '-----------------------------------------
    tbUpperBound.Text = ""
    tbLowerBound.Text = ""
    tbAlpha.Text = ""
    tbEpochMax = ""
    tbSolutionsNA = ""
    tbMovesMin = ""
    tbInput.Text = ""
    x_solution.Text = ""
    y_solution.Text = ""
    tbResults.Text = ""
    tbIterationLog.Text = ""
    imgGraph.visible = False
    '-----------------------------------------
End Sub

Private Sub btnClose_Click()
    'CLOSE BUTTON
    Unload Me
End Sub


Private Sub btnEvaluate_Click()
    
    'VARIABLES
    '-----------------------------------------
     Dim UpperBound As Double           'Random number generator's bounds
     Dim LowerBound As Double
     Dim alpha As Double                'Cooling Parameter
     Dim CurrentFileName As String      'Graph variables
     Dim CurrentChart As Chart
     
     Dim result As SAResult
     '-----------------------------------------
    
    'INPUT ERROR PREVENTION
    '-----------------------------------------
    
    ' Test for empty input
    If tbInput.Text = "" Then
        MsgBox "Please enter a function in the input box"
        Exit Sub
    End If
    
    If tbUpperBound.Text = "" Then
        MsgBox "Please enter an Upper Bound value"
        Exit Sub
    End If
    
    If tbLowerBound.Text = "" Then
        MsgBox "Please enter a Lower Bound value"
        Exit Sub
    End If
    
    If tbEpochMax.Text = "" Then
        MsgBox "Please enter a Max Epochs value"
        Exit Sub
    End If
    
    If tbMovesMin.Text = "" Then
        MsgBox "Please enter a Min Moves value"
        Exit Sub
    End If
    
    If tbSolutionsNA.Text = "" Then
        MsgBox "Please enter an Epochs Without Acceptance value"
        Exit Sub
    End If
    
    If tbAlpha.Text = "" Then
        MsgBox "Please enter an Alpha value"
        Exit Sub
    End If
    
    'Test for Numeric Input
    If Not IsNumeric(tbUpperBound.Text) Then
        MsgBox "Upper Bound must be a number"
        Exit Sub
    End If
    
    If Not IsNumeric(tbLowerBound.Text) Then
        MsgBox "Lower Bound must be a number"
        Exit Sub
    End If
    
    If Not IsNumeric(tbAlpha.Text) Then
        MsgBox "Alpha must be a number"
        Exit Sub
    End If
    
    If Not IsNumeric(tbEpochMax.Text) Then
        MsgBox "Max Epochs must be a number"
        Exit Sub
    End If
    
    If Not IsNumeric(tbMovesMin.Text) Then
        MsgBox "Min Moves must be a number"
        Exit Sub
    End If
    
    If Not IsNumeric(tbSolutionsNA.Text) Then
        MsgBox "Epochs Without Acceptance must be a number"
        Exit Sub
    End If
    
    'Test for valid input conditions
    alpha = tbAlpha.Value
    If alpha <= 0 Or alpha >= 1 Then
        MsgBox "Please enter an Alpha value between 0 and 1 (exclusive)"
        Exit Sub
    End If
    
    If CInt(tbEpochMax.Text) <= 0 Then
        MsgBox "Max Epochs must be a positive number"
        Exit Sub
    End If
    
    If CInt(tbMovesMin.Text) <= 0 Then
        MsgBox "Min Moves must be a positive number"
        Exit Sub
    End If
    
    If CInt(tbSolutionsNA.Text) <= 0 Then
        MsgBox "Epochs Without Acceptance must be a positive number"
        Exit Sub
    End If
    '-----------------------------------------
    
    ' TEST UPPER & LOWER BOUND CONDITIONS
    '-----------------------------------------
    UpperBound = tbUpperBound.Value
    LowerBound = tbLowerBound.Value
    
    If UpperBound - LowerBound < 0 Then      'Test for Upper bound and lower bound conditions
        MsgBox "The upper bound should be larger than the lower bound"
        Exit Sub
    End If
    
    If UpperBound - LowerBound < 1 Then        'Test for Upper bound and lower bound conditions
        MsgBox "The upper and lower bound should be at least 1 apart"
        Exit Sub
    End If
    '-----------------------------------------

    
    ' PREPARE UI
    '-----------------------------------------
    Range("A2:D1000").ClearContents     'Clear excel cells
    Range("A2").Select                  'Set cell A2 as reference
    imgGraph.visible = True             'Make the graph component visible
    '-----------------------------------------
    
    
     'RUN ALGORITHM
    '-----------------------------------------
    result = RunSimulatedAnnealing( _
    tbUpperBound.Value, _
    tbLowerBound.Value, _
    tbAlpha.Value, _
    CInt(tbEpochMax.Text), _
    CInt(tbMovesMin.Text), _
    CInt(tbSolutionsNA.Text), _
    tbInput.Text, _
    opMax.Value)
    '-----------------------------------------
    
    
    'WRITE ITERATION DATA TO SHEET
    '-----------------------------------------
    Range("A2").Select                                  'Start writing from row 2 as row 1 contains headings
    Dim i As Long
    For i = 1 To result.IterationCount
        ActiveCell = result.IterationData(1, i)              'Iteration number column
        ActiveCell.Offset(0, 1) = result.IterationData(2, i) 'Function value column
        ActiveCell.Offset(0, 2) = result.IterationData(3, i) 'y value column
        ActiveCell.Offset(0, 3) = result.IterationData(4, i) 'x value column
        ActiveCell.Offset(1, 0).Select                       'Move to next row / next iteration's data
    Next i
    '-----------------------------------------
    
    
    'DISPLAY RESULTS
    '-----------------------------------------
    tbIterationLog.Text = result.IterationLog
    tbResults.Text = CStr(Format(result.OptimalValue, "0.000"))
    
    If result.HasX Then
        x_solution.Text = CStr(Format(result.x, "0.000"))
    Else
        x_solution.Text = "-"
    End If
    
    If result.HasY Then
        y_solution.Text = CStr(Format(result.y, "0.000"))
    Else
        y_solution.Text = "-"
    End If
    '-----------------------------------------
    
    
    'UPDATE CHART
    '-----------------------------------------
    CurrentFileName = Environ("TEMP") & "\current.gif"
    Set CurrentChart = ThisWorkbook.Sheets("Results").ChartObjects("ConvergenceChart").Chart
    CurrentChart.Export Filename:=CurrentFileName, FilterName:="GIF"
    frmOptimiser.imgGraph.Picture = LoadPicture(CurrentFileName)
    
    MsgBox "Program Completed"
    '-----------------------------------------

End Sub

Private Sub btnHelp_Click() 'Help button display
    MsgBox ("Limitations of the program:" + vbCrLf + vbCrLf + "1) The function must be continuous within the upper and lower bounds" + vbCrLf + vbCrLf + "2) Only the following characters may be used in the function input: +  -  *  /  ^  .  ( )  e  sin( )  cos( )  x  y  rational numbers" + vbCrLf + vbCrLf + "3) Symbols and numbers that are to be multiplied should be seperated by a *. Examples: yx, 3e, 4sinx, 5(3) should be written as y*x, 3*e, 4*sin(x) and 5*(3)" + vbCrLf + vbCrLf + "4) Negative x and y values should be written as: -1*x or -1*y" + vbCrLf + vbCrLf + "5) You can't use a variable to the power of a variable. Example: x^y or y^x" + vbCrLf + vbCrLf + "6) Complicated exponents should be contained in parentheses. Example: x^(3+2/1.5*3)")
End Sub

Private Sub SetAdvancedControlsVisibility(isVisible As Boolean)
    tbMovesMin.visible = isVisible
    tbEpochMax.visible = isVisible
    tbSolutionsNA.visible = isVisible
    tbAlpha.visible = isVisible
    lblMaxIterationsPerEpoch.visible = isVisible
    lblMovesPerEpoch.visible = isVisible
    lblCoolingFactor.visible = isVisible
    lblEpochsWithoutAcceptance.visible = isVisible
End Sub

Private Sub ResetDefaults()
    tbUpperBound.Value = 2
    tbLowerBound.Value = -2
    tbAlpha.Value = 0.95
    tbSolutionsNA.Value = 50
    tbEpochMax.Value = 100
    tbMovesMin.Value = 100
    tbInput.Value = "(x^2+y-11)^2+(x+y^2-7)^2"
End Sub

Private Sub btnReset_Click()
    ResetDefaults
End Sub

Private Sub TabOption_Change()      'Standard preset & Advanced preset
    SetAdvancedControlsVisibility (TabOption.Value <> 0)
    ResetDefaults
End Sub

Private Sub UserForm_Initialize()
    TabOption_Change
End Sub
