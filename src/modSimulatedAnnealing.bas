Attribute VB_Name = "modSimulatedAnnealing"

Public Type SAResult
    x As Double
    y As Double
    OptimalValue As Double
    IterationLog As String
    
    ' Track whether there is a X, Y or both X & Y value in the function
    HasX As Boolean
    HasY As Boolean
    
    ' Track algorithm iterations to plot Convergence Chart in frmOptimiser
    ' Each iteration has 4 values (iterations, function value, y value, x value)
    ' Print data point on chart every 200 iterations
    IterationData() As Variant    ' 2D array of iterations
    IterationCount As Long        ' number of rows collected
End Type

' FUNCTION EVALUATION: get input equation from user and evaluate it based on parameter values
'-----------------------------------------
Public Function EvaluateFunction(x As Double, y As Double, FunctionString As String) As Double
    Dim expr As String
    expr = Replace(FunctionString, "x", x)
    expr = Replace(expr, "y", y)
    expr = Replace(expr, ",", ".")
    expr = Replace(expr, "e", Exp(1))
    EvaluateFunction = Evaluate(expr)
    
End Function

' LOG ROW FORMATTING: builds the standard formatting row for an iteration's log
'-----------------------------------------
Private Function FormatLogRow(Iterations As String, x_Val As String, y_Val As String, func_Val As String) As String
    FormatLogRow = Iterations + Chr(9) + "|" + Chr(9) + _
        x_Val + Chr(9) + Chr(9) + "|" + Chr(9) + _
        y_Val + Chr(9) + Chr(9) + "|" + Chr(9) + _
        func_Val
End Function

Public Function RunSimulatedAnnealing( _
    UpperBound As Double, _
    LowerBound As Double, _
    alpha As Double, _
    Epochs_max As Integer, _
    Moves_min As Integer, _
    Sols_Not_Acc_max As Integer, _
    FunctionString As String, _
    FindMax As Boolean _
) As SAResult
    
    'VARIABLES
    '-----------------------------------------
     Dim Iterations As Double           'Number of iterations
    
     Dim Epochs As Integer              'Counting number of epochs that have occured
     Dim Moves As Integer               'Counting number of moves occured within an epoch
    
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
     
     Dim energy_Diff As Double          'Energy difference between current guess and neighbouring guess
     Dim control_Factor As Double       'Ensures R_comp (comparison value) does not cause an overflow error
     
     Dim R As Double                    'Random number between 0 and 1
     Dim R_comp As Double               'Value R is compared against => algorithm comparison value
     
     Dim dataCount As Long
     dataCount = 0
     
     Dim funcResult As String
     Dim result As SAResult
     '-----------------------------------------
    
    ' VARIABLE INITIATION
    '-----------------------------------------
    Dim logText As String
    
    If FindMax = True Then      'Setup max / min in the simulation table
        logText = FormatLogRow("Iteration:", "X-Values:", "Y-Values:", "Maximum Values:")
    Else
        logText = FormatLogRow("Iteration:", "X-Values:", "Y-Values:", "Minimum Values:")
    End If
    
    'Values are assigned to variables
    Sols_Not_Acc = 0
    control_Factor = 1
    Iterations = 0
    min_Temp = 0.00001
    Temp = 1000
    Repeated_val = 0
    
    Position_y = InStr(1, FunctionString, "y", vbBinaryCompare)       'Finds the position of y in the input
    Position_x = InStr(1, FunctionString, "x", vbBinaryCompare)       'Finds the position of x in the input
        
        
    ' RANDOM NUMBER GENERATION
    '-----------------------------------------
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
            energy_Diff = EvaluateFunction(x_current, y_current, FunctionString) - EvaluateFunction(x_prime, y_prime, FunctionString)
            '-----------------------------------------
            
            
            'SOLUTION ACCEPTANCE DECISION FACTOR
            '-----------------------------------------
            Randomize
            R = Rnd()
            
            If FindMax = True Then                      'Test if the max/min should be calculated
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
                    dataCount = dataCount + 1
                    ReDim Preserve result.IterationData(1 To 4, 1 To dataCount) 'Redefine array to grow it by one column
                    
                    result.IterationData(1, dataCount) = Iterations
                    result.IterationData(2, dataCount) = EvaluateFunction(x_current, y_current, FunctionString)
                    result.IterationData(3, dataCount) = y_current
                    result.IterationData(4, dataCount) = x_current
                    
                    If Position_x = 0 Then                  'Tests if x is present in input
                        logText = logText & vbCrLf & FormatLogRow(CStr(Iterations), "-", CStr(Format(y_current, "0.000")), CStr(Format((EvaluateFunction(x_current, y_current, FunctionString)), "0.000")))
                    ElseIf Position_y = 0 Then                  'Tests if y is present in input
                        logText = logText & vbCrLf & FormatLogRow(CStr(Iterations), CStr(Format(x_current, "0.000")), "-", CStr(Format((EvaluateFunction(x_current, y_current, FunctionString)), "0.000")))
                    ElseIf (Position_x > 0) And (Position_y > 0) Then  'Tests if x and y is present in input
                        logText = logText & vbCrLf & FormatLogRow(CStr(Iterations), CStr(Format(x_current, "0.000")), CStr(Format(y_current, "0.000")), CStr(Format((EvaluateFunction(x_current, y_current, FunctionString)), "0.000")))
                    End If
                
                End If
                 '-----------------------------------------
            End If
            
            Epochs = Epochs + 1                     'Update number of moves within an epoch
            If Temp < 50 Then                       'Repeated value algorithm
             If CStr(Format(EvaluateFunction(x_prev, y_prev, FunctionString), "0.00000")) = CStr(Format(EvaluateFunction(x_current, y_current, FunctionString), "0.00000")) Then
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


   'APPEND FINAL RESULT TO LOG
    '-----------------------------------------
    funcResult = CStr(Format(EvaluateFunction(x_current, y_current, FunctionString), "0.000"))
    
    If (Position_x = 0) And (Position_y = 0) Then
        logText = logText & vbCrLf & FormatLogRow("1", "-", "-", funcResult)
    ElseIf (Position_x > 0) And (Position_y > 0) Then
        logText = logText & vbCrLf & FormatLogRow(CStr(Iterations), CStr(Format(x_current, "0.000")), CStr(Format(y_current, "0.000")), funcResult)
    ElseIf Position_x = 0 Then
        logText = logText & vbCrLf & FormatLogRow(CStr(Iterations), "-", CStr(Format(y_current, "0.000")), funcResult)
    ElseIf Position_y = 0 Then
        logText = logText & vbCrLf & FormatLogRow(CStr(Iterations), CStr(Format(x_current, "0.000")), "-", funcResult)
    End If
    '-----------------------------------------

    ' RETURN ALGORITHM RESULTS
    '-----------------------------------------
    result.x = x_current
    result.y = y_current
    result.HasX = (Position_x > 0)
    result.HasY = (Position_y > 0)
    result.OptimalValue = EvaluateFunction(x_current, y_current, FunctionString)
    result.IterationLog = logText
    result.IterationCount = dataCount
    
    RunSimulatedAnnealing = result
End Function
