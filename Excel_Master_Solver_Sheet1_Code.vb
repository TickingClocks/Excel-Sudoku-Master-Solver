Option Explicit
'clears cell contents
Private Sub clear_Click()
    
    Dim k As Integer
    Dim j As Integer
    
    j = 1
    
    Do While (j <= 9)
        
        k = 2
        Do While (k <= 10)
            
            With Range(Cells(k, j), Cells(k, j))
                .Value = ""
                .Font.Color = RGB(73, 149, 200)
                .Font.Size = 30
                .Font.Bold = True
                .Interior.Color = RGB(237, 237, 237)
            End With
        
            k = k + 1
            
        Loop
        
        j = j + 1
        
    Loop
    
End Sub
'disable pass messages
Private Sub notOkMsg_Click()
    
    If Sheets(1).notOkMsg Then
    
        Sheets(1).okMsg.Value = False
        
    Else
    
        Sheets(1).okMsg.Value = True
        
    End If
    
End Sub
'enable pass messages
Private Sub okMsg_Click()
    
    If Sheets(1).okMsg Then
    
        Sheets(1).notOkMsg.Value = False
        
    Else
    
        Sheets(1).notOkMsg.Value = True
        
    End If
    
End Sub
'one pass
Private Sub onePass_Click()
    
    If Sheets(1).onePass Then
        
        Sheets(1).pass100.Value = False
        
    Else
    
        Sheets(1).pass100.Value = True
        
    End If
    
End Sub
'100 passes
Private Sub pass100_Click()
        
    If Sheets(1).pass100 Then
        
        Sheets(1).onePass.Value = False
        
    Else
    
        Sheets(1).onePass.Value = True
        
    End If
        
End Sub
'searches for solution to sudoku puzzle
Private Sub solve_Click()
    
    Dim eCheck As String
    
    Dim cellCheck As Integer 'used to check for blank spaces
    Dim try As Integer 'current try value
    Dim entry As Integer 'final entry
    Dim xCheck As Integer
    Dim yCheck As Integer
    Dim sqCheck As Integer
    Dim k As Integer 'used for CURRENT POSITION column
    Dim j As Integer 'used for CURRENT POSITION row
    Dim D As Integer 'used for CHECKING Rows/Columns
    Dim a As Integer 'used for CHECKING square
    Dim b As Integer 'used for CHECKING square
    Dim c As Integer 'used for CHECKING square (loop)
    Dim E As Integer 'used for CHECKING square (loop)
    Dim beSure As Integer 'used to ensure placement is the only valid option for cell
    Dim genMsg As Integer
    Dim z As Long
    Dim p As Long
    
    Dim onePass As Boolean
    Dim pass100 As Boolean
    Dim FAIL As Boolean
    Dim skip As Boolean
    Dim okMsg As Boolean
    
    pass100 = Sheets(1).pass100.Value
    onePass = Sheets(1).onePass.Value
    okMsg = Sheets(1).okMsg.Value
    FAIL = False
    skip = False
    
    If (pass100 = True) Then 'go through puzzle 100 times
    
        z = 100
         
    End If
    
    If (onePass = True) Then 'go through puzzle 1 time
    
        z = 1
        
    End If
    
    If (onePass <> True) And (pass100 <> True) Then 'no selection revert to 1 pass
        
        genMsg = MsgBox("Pass selection not made. Please select the number of passes and try again.", vbOKOnly, "SELECT A PASS SETTING!")
        FAIL = True
        
    End If
    
    If (FAIL <> True) Then 'good to go!
        p = 1
        Do While (p <= z) 'loop for board iteration
            
            k = 1
            Do While (k <= 9) 'moving over one column
                
                j = 2
                Do While (j <= 10) 'moving down one row
                    
                    cellCheck = CInt(Range(Cells(j, k), Cells(j, k)))
                    
                    If (cellCheck = 0) Then 'empty space
                        
                        'reset variables
                        beSure = 0
                        try = 1
                        Do While (try <= 9) 'check each possible number 1-9
                            
                            ''''''''''''''''''''''''''''''''''''''''''''''''
                            '''''''''''check if valid in row''''''''''''''''
                            ''''''''''''''''''''''''''''''''''''''''''''''''
                            skip = False
                            D = 1
                            Do While (D <= 9)
                            
                                If (D = k) Then 'current 'empty' cell. DONT check
                                
                                    D = D + 1 'skip past
                                    
                                Else
    
                                    xCheck = CInt(Range(Cells(j, D), Cells(j, D)))
                                    
                                    'check if = try
                                    If (xCheck = try) Then 'found try in row - INVALID ENTRY
                                    
                                        skip = True 'dont bother checking column and square
                                        D = 10 'get out of d loop
                                        
                                    Else
                                        
                                        D = D + 1
                                    
                                    End If
                                    
                                End If
                                
                            Loop
                            
                            ''''''''''''''''''''''''''''''''''''''''''''''''
                            ''''''''''''check if valid in column''''''''''''
                            ''''''''''''''''''''''''''''''''''''''''''''''''
                            
                            If (skip <> True) Then
                            
                                D = 2 'reset d
                                Do While (D <= 10)
                                
                                    If (D = j) Then 'current 'empty cell. DONT check
                                    
                                        D = D + 1
                                        
                                    Else
                                    
                                        yCheck = CInt(Range(Cells(D, k), Cells(D, k)))
                                        
                                        If (yCheck = try) Then 'found try in column - INVALID Entry
                                            
                                            skip = True 'dont bother checking square
                                            D = 11 'get out of d loop
                                        
                                        Else
                                            
                                            D = D + 1
                                                                            
                                        End If
                                        
                                    End If
                                    
                                Loop
                            
                            End If
                            
                            ''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''''''''''check if valid in square''''''''''''''''
                            ''''''''''''''''''''''''''''''''''''''''''''''''''
                            
                            If (skip <> True) Then
								
								'feed in x and y coordinates and get top left coordinate for the current 3x3 block
                                a = 3 * Int((j - 2) / 3) + 2
                                b = 3 * Int((k - 1) / 3) + 1
                                
                                c = 0
                                Do While (c <= 2)
                                    
                                    E = 0
                                    Do While (E <= 2)
                                    
                                        sqCheck = CInt(Range(Cells(a, b), Cells(a, b)))
                                        
                                        If (sqCheck = try) Then 'found in square - INVALID Entry
                                        
                                            skip = True
                                            E = 3
                                            c = 3
                                            
                                        Else
                                        
                                            E = E + 1
                                            b = b + 1
                                            
                                            If (E = 3) Then 'reset b
                                            
                                                b = 3 * Int((k - 1) / 3) + 1
                                                
                                            End If
                                            
                                        End If
                                        
                                    Loop
                                    
                                    c = c + 1
                                    a = a + 1
                                    
                                Loop
                                
                            End If
                            
                            ''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''''''''''''after checks''''''''''''''''''''''''''
                            ''''''''''''''''''''''''''''''''''''''''''''''''''
                            
                            If (skip <> True) Then 'try wasnt skipped
                                
                                beSure = beSure + 1
                                entry = try
                                
                            End If
                            
                            try = try + 1
                            
                        Loop
                        
                        'POPULATE CELL
                        If (beSure = 1) Then 'found only one option
                        
                            Range(Cells(j, k), Cells(j, k)).Value = entry
                            Range(Cells(j, k), Cells(j, k)).Font.Color = RGB(190, 85, 85)
                            Range(Cells(j, k), Cells(j, k)).Font.Bold = True
                            
                        End If
                        
                        j = j + 1
    
                    Else
                    
                        j = j + 1
                    
                    End If
                    
                Loop
                
                k = k + 1
                
            Loop 'close loop k
        
            p = p + 1
            
        Loop 'close loop p
        
        If (okMsg = True) Then
            
            'pass message
            MsgBox ("Finished " + CStr(p - 1) + " iterations!")
            
        End If
        
    End If
     
End Sub
'fill empty cells with 0
Private Sub zeros_Click()
    
    Dim check As String
    
    Dim k As Integer
    Dim j As Integer
    
    j = 1
    
    Do While (j <= 9)
        
        k = 2
        Do While (k <= 10)
            
            check = CStr(Range(Cells(k, j), Cells(k, j)).Value)
            
            If (check = "") Then 'blank cell
            
                Range(Cells(k, j), Cells(k, j)).Value = 0
                Range(Cells(k, j), Cells(k, j)).Font.Color = vbRed
                
            End If
        
            k = k + 1
            
        Loop
        
        j = j + 1
        
    Loop
    
End Sub
