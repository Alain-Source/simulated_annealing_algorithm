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
imgGraph.Visible = False
'-----------------------------------------

End Sub
Private Sub btnClose_Click()
'CLOSE BUTTON
End
End Sub


Private Sub btnEvaluate_Click()

'VARIABLES
'-----------------------------------------
 Dim Iterations As Double           'Number of iterations

 Dim UpperBound As Double           'Random number generator's bounds
 Dim LowerBound As Double
  
 Dim Epochs_max As Integer          'Maximum number of epochs
 Dim Moves_min As Integer           'Minimum number of moves per epoch
 Dim Epochs As Integer              'Counting number of epochs that have occured
 Dim Moves As Integer               'Counting number of moves occured within an epoch

 Dim Sols_Not_Acc_max As Integer    'Maximum epochs that may pass without acceptance of a new solution
 Dim Sols_Not_Acc As Integer        'Epochs that may pass without acceptance of a new solution
 
 Dim x_prime As Double              'Neighbouring guess for x
 Dim x_current As Double            'Current x value
 Dim x_prev As Double               'Saves previous x value when x_current changes
 
 Dim y_prime As Double              'Neighbouring guess for y
 Dim y_current As Double            'Current y value
 Dim y_prev As Double               'Saves previous y value when x_current changes
 
 Dim Repeated_val As Double         'Counts how many times the max remains the same
 
 Dim Position_y As Integer          'Finds the position of y in the input string
 Dim Position_x As Integer          'Finds the position of x in the input string
 
 Dim Temp As Double                 'Algorithm Temperature
 Dim min_Temp As Double             'Minimum Temperature
 Dim alpha As Double                'Cooling Parameter
 
 Dim energy_Diff As Double          'Energy difference between current guess and neighbouring guess
 Dim control_Factor As Double       'Ensures R_comp (comparison value) does not cause an overflow error
 
 Dim R As Double                    'Random number between 0 and 1
 Dim R_comp As Double               'Value R is compared against => algorithim comparison value
 
 Dim CurrentFileName As String      'Graph variables
 Dim CurrentChart As Chart
 '-----------------------------------------




'INPUT ERROR PREVENTION
'-----------------------------------------
If tbInput.Text = "" Then           'Test for empty function box
    MsgBox "Please enter a function in the input box"
    Exit Sub
End If

If tbUpperBound.Text = "" Then       'Test for empty upper bound box
    MsgBox "Please enter an Upper Bound value"
    Exit Sub
End If

If tbLowerBound.Text = "" Then       'Test for empty lower bound box
    MsgBox "Please enter a Lower Bound value"
    Exit Sub
End If

If tbAlpha.Text = "" Then            'Test for empty alpha box
    MsgBox "Please enter an Alpha value"
    Exit Sub
End If

If tbEpochMax.Text = "" Then           'Test for empty epoch box
    MsgBox "Please enter a Max Epochs value"
    Exit Sub
End If

If tbMovesMin.Text = "" Then           'Test for empty moves box
    MsgBox "Please enter a Min Moves value"
    Exit Sub
End If

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

alpha = tbAlpha.Value
If alpha >= 1 Then                       'Test for alpha box conditions
    MsgBox "Please enter an Alpha value < 1"
    Exit Sub
End If
'-----------------------------------------




' VARIABLE INITIATION
'-----------------------------------------
If opMax.Value = True Then      'Setup max / min in the simulation table
tbIterationLog.Text = "Iteration:" + Chr(9) + "|" + Chr(9) + "X-Values:" + Chr(9) + Chr(9) + "|" + Chr(9) + "Y-Values:" + Chr(9) + Chr(9) + "|" + Chr(9) + "Maximum Values:"
Else
tbIterationLog.Text = "Iteration:" + Chr(9) + "|" + Chr(9) + "X-Values:" + Chr(9) + Chr(9) + "|" + Chr(9) + "Y-Values:" + Chr(9) + Chr(9) + "|" + Chr(9) + "Minimum Values:"
End If

'Values are assigned to variables
Sols_Not_Acc = 0
Sols_Not_Acc_max = tbSolutionsNA.Text
Moves_min = tbMovesMin.Text
Epochs_max = tbEpochMax.Text
control_Factor = 1
Iterations = 0
min_Temp = 0.00001
Temp = 1000
Repeated_val = 0

Position_y = InStr(1, tbInput.Text, "y", vbBinaryCompare)       'Finds the position of y in the input
Position_x = InStr(1, tbInput.Text, "x", vbBinaryCompare)       'Finds the position of x in the input

Range("A2:D1000").ClearContents     'Clear excel cells
Range("A2").Select                  'Set cell A2 as reference
imgGraph.Visible = True             'Make the graph component visible
'-----------------------------------------



' RANDOM NUMBER GENERATION
'-----------------------------------------
UpperBound = tbUpperBound.Value
LowerBound = tbLowerBound.Value

Randomize
x_current = Int((UpperBound - LowerBound) * Rnd() + LowerBound) 'Generates random numbers between the upper and lower bounds
y_current = Int((UpperBound - LowerBound) * Rnd() + LowerBound)
'-----------------------------------------



' SIMULATED ANNEALING LOOP
While (Sols_Not_Acc <= Sols_Not_Acc_max) And (Repeated_val < 900) And (Position_y > 0 Or Position_x > 0)        'Outer while loop
    Moves = 0 'Counting number of moves occured within an epoch
    Epochs = 0 'Counting number of epochs that have occured


    While ((Epochs <= Epochs_max) And (Moves < Moves_min) And (Temp > min_Temp))
    
    
    
        'GENERATE RANDOM NEIGHBOURING SOLUTION
        '-----------------------------------------
        If ((x_current - 0.5) < LowerBound) Then        'If current x value is within 0.5 of the lowerbound:
            Randomize
            If Rnd() < 0.5 Then
                x_prime = x_current - ((x_current - LowerBound) * Rnd())
                Else
                x_prime = x_current + (Rnd() * 0.5)
            End If
        End If
        
        
        If ((x_current + 0.5) > UpperBound) Then        'If current x value is within 0.5 of the upperbound:
            Randomize
            If Rnd() < 0.5 Then
                x_prime = x_current + (UpperBound - x_current) * Rnd()
                Else
                x_prime = x_current - (Rnd() * 0.5)
            End If
        End If
        
        If ((x_current - 0.5) >= LowerBound) And ((x_current + 0.5) <= UpperBound) Then   'If current x value is not within 0.5 of bounds:
            x_prime = x_current + (Rnd() - 0.5)
        End If
        
        If ((y_current - 0.5) < LowerBound) Then    'If current y value is within 0.5 of the lowerbound:
            Randomize
            If Rnd() > 0.5 Then
                y_prime = y_current - ((y_current - LowerBound) * Rnd())
                Else
                y_prime = y_current + (Rnd() * 0.5)
            End If
        End If
        
        
        If ((y_current + 0.5) > UpperBound) Then    'If current y value is within 0.5 of the upperbound:
            Randomize
            If Rnd() > 0.5 Then
                y_prime = y_current + (UpperBound - y_current) * Rnd()
                Else
                y_prime = y_current - (Rnd() * 0.5)
            End If
        End If
        
        If ((y_current - 0.5) >= LowerBound) And ((y_current + 0.5) <= UpperBound) Then   'If current y value is not within 0.5 of bounds:
            y_prime = y_current + (Rnd() - 0.5)
        End If
        '-----------------------------------------
        
        
        'ASSESS ENERGY DIFFERENCE: between current & neighbouring solution
        '-----------------------------------------
        energy_Diff = EvaluateFunction(x_current, y_current) - EvaluateFunction(x_prime, y_prime)
        '-----------------------------------------
        
        
        'SOLUTION ACCEPTANCE DECISION FACTOR
        '-----------------------------------------
        Randomize
        R = Rnd()
        
        If opMax.Value = True Then                      'Test if the max/min should be calculated
        control_Factor = -1 * energy_Diff / Temp        'Calculates Maximum
        Else
        control_Factor = 1 * energy_Diff / Temp         'Calculates Minimum
        End If
        
        If (control_Factor > 230) Then                  'Prevents overflow error
          control_Factor = 230
        End If
        
        If 1 <= Exp(control_Factor) Then                'Determine value to compare against for new value to be accepted
            R_comp = 1
        Else
            R_comp = Exp(control_Factor)
        End If
        
        
     
        If R < R_comp Then
            x_current = x_prime                         'Accept new solution
            y_current = y_prime
            Moves = Moves + 1                           'Update number of accepted solutions and iterations
            Iterations = Iterations + 1
        '-----------------------------------------
            
            
            
        'SIMULATION DISPLAY
        '-----------------------------------------
            If (Iterations Mod 200 = 0) Then        'Shows every 200 iterations
                If Position_x = 0 Then                  'Tests if x is present in input
                    tbIterationLog.Text = tbIterationLog.Text & vbCrLf & CStr(Iterations) + Chr(9) + "|" + Chr(9) + "-" + Chr(9) + Chr(9) + "|" + Chr(9) + CStr(Format(y_current, "0.000")) + Chr(9) + Chr(9) + "|" + Chr(9) + CStr(Format((EvaluateFunction(x_current, y_current)), "0.000"))
                End If
                
                If Position_y = 0 Then                  'Tests if y is present in input
                    tbIterationLog.Text = tbIterationLog.Text & vbCrLf & CStr(Iterations) + Chr(9) + "|" + Chr(9) + CStr(Format(x_current, "0.000")) + Chr(9) + Chr(9) + "|" + Chr(9) + "-" + Chr(9) + Chr(9) + "|" + Chr(9) + CStr(Format((EvaluateFunction(x_current, y_current)), "0.000"))
                End If
                
                If (Position_x > 0) And (Position_y > 0) Then   'Tests if x and y is present in input
                    tbIterationLog.Text = tbIterationLog.Text & vbCrLf & CStr(Iterations) + Chr(9) + "|" + Chr(9) + CStr(Format(x_current, "0.000")) + Chr(9) + Chr(9) + "|" + Chr(9) + CStr(Format(y_current, "0.000")) + Chr(9) + Chr(9) + "|" + Chr(9) + CStr(Format((EvaluateFunction(x_current, y_current)), "0.000"))
                End If
                
                ActiveCell = Iterations                 'Puts the algorithm data in excel
                ActiveCell.Offset(0, 1) = EvaluateFunction(x_current, y_current)
                ActiveCell.Offset(0, 2) = y_current
                ActiveCell.Offset(0, 3) = x_current
                ActiveCell.Offset(1, 0).Select
            End If
            
            
        End If
        '-----------------------------------------
        
        
        Epochs = Epochs + 1                     'Update number of moves within an epoch
        If Temp < 50 Then                       'Repeated value algorithm
         If CStr(Format(EvaluateFunction(x_prev, y_prev), "0.00000")) = CStr(Format(EvaluateFunction(x_current, y_current), "0.00000")) Then
            Repeated_val = Repeated_val + 1
            Else
            Repeated_val = 0
         End If
        End If
            x_prev = x_current
            y_prev = y_current
    Wend 'End Inner WHILE loop
    
    
    
    'UPDATE TEMPERATURE & INCREMENT EPOCHS
    '-----------------------------------------
    Temp = alpha * Temp                         'Update temperature
    If Moves = 0 Then                           'If number of moves (in epoch) = 0 do:
        Sols_Not_Acc = Sols_Not_Acc + 1         'Increment number of epochs
    End If
    '-----------------------------------------
    
    
Wend 'End Outer WHILE loop


'OUTPUT FINAL SOLUTION
'-----------------------------------------
If Position_y = 0 Then                                  'Test if y is present in input
    x_solution.Text = CStr(Format(x_current, "0.000"))
    y_solution.Text = "-"
End If

If Position_x = 0 Then                                  'Test if x is present in input
    y_solution.Text = CStr(Format(y_current, "0.000"))
    x_solution.Text = "-"
End If

If (Position_x = 0) And (Position_y = 0) Then           'Test if x and y is not present in input
    y_solution.Text = "-"
    x_solution.Text = "-"
End If

If (Position_y > 0) And (Position_x > 0) Then           'Test if x and y is present in input
    y_solution.Text = CStr(Format(y_current, "0.000"))
    x_solution.Text = CStr(Format(x_current, "0.000"))
End If

If x_solution.Text = "0.000" Then
    x_solution.Text = CStr(x_current)
End If

If y_solution.Text = "0.000" Then
    y_solution.Text = CStr(y_current)
End If

tbResults.Text = CStr(Format((EvaluateFunction(x_current, y_current)), "0.000"))    'Display Max/Min


CurrentFileName = Environ("TEMP") & "\current.gif"                             'Upload graph to UI
Set CurrentChart = ThisWorkbook.Sheets("Results").ChartObjects("ConvergenceChart").Chart
CurrentChart.Export Filename:=CurrentFileName, FilterName:="GIF"
frmOptimiser.imgGraph.Picture = LoadPicture(CurrentFileName)


If (Position_x = 0) And (Position_y = 0) Then          'Display max/min if there are no x or y values in input
    tbIterationLog.Text = tbIterationLog.Text & vbCrLf & "1" + Chr(9) + "|" + Chr(9) + "-" + Chr(9) + Chr(9) + "|" + Chr(9) + "-" + Chr(9) + Chr(9) + "|" + Chr(9) + CStr(Format((EvaluateFunction(x_current, y_current)), "0.000"))
    MsgBox "Program Completed"
    Exit Sub
End If

If (Position_x > 0) And (Position_y > 0) Then       'If there are x and y values in input display answer
    tbIterationLog.Text = tbIterationLog.Text & vbCrLf & CStr(Iterations) + Chr(9) + "|" + Chr(9) + CStr(Format(x_current, "0.000")) + Chr(9) + Chr(9) + "|" + Chr(9) + CStr(Format(y_current, "0.000")) + Chr(9) + Chr(9) + "|" + Chr(9) + CStr(Format((EvaluateFunction(x_current, y_current)), "0.000"))
End If

If Position_x = 0 Then      'If there are only y values in input display answer
    tbIterationLog.Text = tbIterationLog.Text & vbCrLf & CStr(Iterations) + Chr(9) + "|" + Chr(9) + "-" + Chr(9) + Chr(9) + "|" + Chr(9) + CStr(Format(y_current, "0.000")) + Chr(9) + Chr(9) + "|" + Chr(9) + CStr(Format((EvaluateFunction(x_current, y_current)), "0.000"))
End If

If Position_y = 0 Then      'If there are only x values in input display answer
    tbIterationLog.Text = tbIterationLog.Text & vbCrLf & CStr(Iterations) + Chr(9) + "|" + Chr(9) + CStr(Format(x_current, "0.000")) + Chr(9) + Chr(9) + "|" + Chr(9) + "-" + Chr(9) + Chr(9) + "|" + Chr(9) + CStr(Format((EvaluateFunction(x_current, y_current)), "0.000"))
End If

MsgBox "Program Completed"
'-----------------------------------------
End Sub


' FUNCTION EVALUATION: get input equation from user and evaluate it based on parameter values
'-----------------------------------------
Public Function EvaluateFunction(X As Double, Y As Double) As Double
Dim expr As String
expr = Replace(tbInput.Value, "x", X)
expr = Replace(expr, "y", Y)
expr = Replace(expr, ",", ".")
expr = Replace(expr, "e", Exp(1))
EvaluateFunction = Evaluate(expr)
'-----------------------------------------
End Function


Private Sub btnHelp_Click() 'Help button display
MsgBox ("Limitations of the program:" + vbCrLf + vbCrLf + "1) The function must be continuous within the upper and lower bounds" + vbCrLf + vbCrLf + "2) Only the following characters may be used in the function input: +  -  *  /  ^  .  ( )  e  sin( )  cos( )  x  y  rational numbers" + vbCrLf + vbCrLf + "3) Symbols and numbers that are to be multiplied should be seperated by a *. Examples: yx, 3e, 4sinx, 5(3) should be written as y*x, 3*e, 4*sin(x) and 5*(3)" + vbCrLf + vbCrLf + "4) Negative x and y values should be written as: -1*x or -1*y" + vbCrLf + vbCrLf + "5) You can't use a variable to the power of a variable. Example: x^y or y^x" + vbCrLf + vbCrLf + "6) Complicated exponents should be contained in parentheses. Example: x^(3+2/1.5*3)")
End Sub

Private Sub btnReset_Click()        'Reset button
tbUpperBound.Value = 2
tbLowerBound.Value = -2
tbAlpha.Value = 0.95
tbSolutionsNA.Value = 50
tbEpochMax.Value = 100
tbMovesMin.Value = 100
tbInput.Value = "(x^2+y-11)^2+(x+y^2-7)^2"
End Sub

Private Sub lblEpochsWithoutAcceptance_Click()

End Sub

Private Sub lblMovesPerEpoch_Click()

End Sub

Private Sub TabOption_Change()      'Standard preset and Advanced preset
If TabOption.Value = 0 Then
tbMovesMin.Visible = False
tbEpochMax.Visible = False
tbSolutionsNA.Visible = False
tbAlpha.Visible = False
lblMaxIterationsPerEpoch.Visible = False
lblMovesPerEpoch.Visible = False
lblCoolingFactor.Visible = False
lblEpochsWithoutAcceptance.Visible = False
Else
tbMovesMin.Visible = True
tbEpochMax.Visible = True
lblMaxIterationsPerEpoch.Visible = True
tbAlpha.Visible = True
tbSolutionsNA.Visible = True
lblMovesPerEpoch.Visible = True
lblCoolingFactor.Visible = True
lblEpochsWithoutAcceptance.Visible = True
End If
tbUpperBound.Value = 2
tbLowerBound.Value = -2
tbAlpha.Value = 0.95
tbSolutionsNA.Value = 50
tbEpochMax.Value = 100
tbMovesMin.Value = 100
tbInput.Value = "(x^2+y-11)^2+(x+y^2-7)^2"
End Sub

Private Sub tbAlpha_Change()

End Sub

Private Sub tbResults_Change()

End Sub

Private Sub UserForm_Initialize()
    TabOption_Change
End Sub
