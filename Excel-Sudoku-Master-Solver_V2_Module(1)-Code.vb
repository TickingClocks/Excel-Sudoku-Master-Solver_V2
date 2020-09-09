Option Explicit
'Global Variables
Public FAIL As Boolean
Public genMsg As Integer
Public valMsg As Boolean
'Main Function
Public Function runSolve()

   FAIL = False
   
   ''''''''''''''''''''''''
   '''Ensure valid Board'''
   ''''''''''''''''''''''''
    
    'doesnt show valid message
    valMsg = False
    
    Call ensureValid
    
    
    ''''''''''''''''''''''''
    ''''find solution'''''''
    ''''''''''''''''''''''''
    
    Call solvePuzzle
    
    '''''''''''''''''''''''''

End Function
'Clears Spreadsheet
Public Function clearSheet()
    
    Dim x As Integer
    Dim y As Integer
    
    x = 1
    Do While (x <= 9)
        
        y = 2
        Do While (y <= 10)
        
            With Cells(y, x)
                .Value = ""
                .Font.Name = "Georgia"
                .Font.Size = 36
                .Font.Bold = True
                .Font.Color = RGB(52, 131, 202)
                .Interior.Color = RGB(234, 220, 244)
                .VerticalAlignment = xlVAlignBottom
                
            End With
            
            y = y + 1
            
        Loop
        
        x = x + 1
        
    Loop
    
End Function
