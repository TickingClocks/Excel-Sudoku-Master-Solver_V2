''''''''''''''''''''''''''''''
''''''''SHEET 1 CODE''''''''''
''''''''''''''''''''''''''''''

Option Explicit
'Clears Spreadsheet
Public Sub clear_Click()
        
    Call clearSheet
    
End Sub

'''''''''''''''''''''''''''''''''''''''''

Private Sub CommandButton1_Click()

    valMsg = True
    
    Call ensureValid
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''

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

