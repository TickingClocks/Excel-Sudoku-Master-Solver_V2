Option Explicit
'Clear button
Public Sub clear_Click()
        
    Call clearSheet
    
End Sub
'Validate Button
Private Sub CommandButton1_Click()

    Dim k As Integer
    Dim j As Integer

    '''''reset font size and color''''
    
    k = 2
    Do While (k <= 10)
        
        j = 1
        Do While (j <= 9)
        
            Cells(k, j).Font.Size = 36
            Cells(k, j).Font.Color = RGB(52, 131, 202)
            
            j = j + 1
            
        Loop
        
        k = k + 1
        
    Loop
    
    'display valid board message
    valMsg = True
    
    'check if board is valid
    Call ensureValid
    
End Sub
'SLO-MO button
Private Sub sloMo_Click()

    If (Range(Cells(6, 10), Cells(6, 14)).Interior.Color = vbGreen) Then 'change red
    
        Range(Cells(6, 10), Cells(6, 14)).Interior.Color = vbRed
        Range("O11").Interior.Color = vbRed
        
    Else 'change green
        
        Range(Cells(6, 10), Cells(6, 14)).Interior.Color = vbGreen
        Range("O11").Interior.Color = vbGreen
        
    End If

    
End Sub
'Solves the Sudoku
Private Sub solve_Click()

    Dim j As Integer
    Dim k As Integer
    
    '''''reset font size to ensure backtracking works''''
    
    k = 2
    Do While (k <= 10)
        
        j = 1
        Do While (j <= 9)
        
            Cells(k, j).Font.Size = 36
            Cells(k, j).Font.Color = RGB(52, 131, 202)
            
            j = j + 1
            
        Loop
        
        k = k + 1
        
    Loop
    
    'call main function
    Call runSolve
    
End Sub


