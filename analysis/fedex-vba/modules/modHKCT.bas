Attribute VB_Name = "modHKCT"
'Option Explicit

Dim objDataRange As Excel.Range '- An Excel cell range object - required
Dim intRowCount As Integer      '- Counts total # of rows
Dim intColCount As Integer      '- Counts total # of columns
Dim intCounter As Integer       '- An interative counter for rows
Dim intVCounter As Integer      '- An interative counter for columns
Dim strPath As String           '- Global indicator of the save path

'Event for clicking the form button
Sub Process_Click()

'***************Initialise************************
'Set the worksheet named 'Trans.' as the active sheet
Trans.Activate
Set objDataRange = Trans.UsedRange
intRowCount = objDataRange.Rows.Count
intColCount = objDataRange.Columns.Count
Range("A1").Select
Dim x As Variant


'**************'Qualify if there are any rows in the Excel****************'
    'Only work if there are more than 2 rows - as the 1st two rows are titles
    If intRowCount > 2 Then
'**************Open a text file for printing******************************
        'A drive selection box appears to select save location
        Dim vblnPathExists As Boolean: vblnPathExists = False
        'Set path separately
        strPath = FileHandler(vblnPathExists)
        
        If vblnPathExists = True Then
            Open strPath For Output Access Write As #1
            Range("A1").Select
                'This loop - loops for the # of rows
                For intCounter = 4 To intRowCount
                With ActiveCell
                    'This second loops each column for a particular row
                    If Trim(CStr(.Cells(intCounter, 1).Value)) <> "" And Trim(CStr(.Cells(intCounter, 1).Value)) <> Trim(CStr(.Cells(intCounter - 1, 1).Value)) Then
                        
                        Print #1, "0," & Chr(34) & "020" & Chr(34)
                        For i = 1 To Default.UsedRange.Rows.Count
                            If Trim(CStr(Default.Cells(i, 3).Value)) <> "" Then
                                Print #1, CStr(Default.Cells(i, 2).Value) & "," & Chr(34) & CStr(Default.Cells(i, 3).Value) & Chr(34)
                            End If
                        Next i
                        For intVCounter = 1 To intColCount
                            If Strings.Trim(.Cells(3, intVCounter).Value) <> "" And Strings.Trim(.Cells(intCounter, intVCounter).Value) <> "" Then
                                Debug.Print .Cells(3, intVCounter).Value
                                Debug.Print .Cells(intCounter, intVCounter).Value
                                
                                CommodCount = 1
                                                
                                If .Cells(2, intVCounter).Value <> "" And IsNumeric(.Cells(intCounter, intVCounter).Value) Then
                                    Print #1, CStr(.Cells(3, intVCounter).Value) & "," & Chr(34) & CStr(.Cells(intCounter, intVCounter).Value * CDec(.Cells(2, intVCounter).Value)) & Chr(34)
                                Else
                                    Print #1, CStr(.Cells(3, intVCounter).Value) & "," & Chr(34) & CStr(.Cells(intCounter, intVCounter).Value) & Chr(34)
                                End If
                                                
                                        
                              
                            End If
                            
                        'Move the column counter one place to the right and loop again vertically
                        Next intVCounter
                    ElseIf Trim(CStr(.Cells(intCounter, 1).Value)) = Trim(CStr(.Cells(intCounter - 1, 1).Value)) Then
                    CommodCount = CommodCount + 1
                        For intVCounter = 1 To intColCount
                        
                        x = Application.VLookup(CStr(.Cells(3, intVCounter).Value), GroupFields.Range("A:A"), 1, False)
                        If IsError(x) Then
                            Else
                                If .Cells(2, intVCounter).Value <> "" And IsNumeric(.Cells(intCounter, intVCounter).Value) Then
                                    Print #1, CStr(.Cells(3, intVCounter).Value) & "-" & CommodCount & "," & Chr(34) & CStr(.Cells(intCounter, intVCounter).Value * CDec(.Cells(2, intVCounter).Value)) & Chr(34)
                                Else
                                    '(2022-12-07 Logan) Check If cell is empty (null string) This is to avoid printing empty value in the file
                                    If Trim(CStr(.Cells(intCounter, intVCounter).Value) & vbNullString) <> vbNullString Then
                                        Print #1, CStr(.Cells(3, intVCounter).Value) & "-" & CommodCount & "," & Chr(34) & CStr(.Cells(intCounter, intVCounter).Value) & Chr(34)
                                    End If
                                End If
                                   
                            End If
                        Next intVCounter
                        
                    End If
                'Move the row counter one place down and loop again horizontally
                
                
                If Trim(CStr(.Cells(intCounter, 1).Value)) <> "" And Trim(CStr(.Cells(intCounter, 1).Value)) <> Trim(CStr(.Cells(intCounter + 1, 1).Value)) Then
                    Print #1, "99," & Chr(34) & vbNullString & Chr(34)
                End If
                End With
                Next intCounter
                
            'Close the text file and outputs content
            Close #1
        End If
    End If
End Sub


'******* Function to handle opening and closing of the text file *************
Private Function FileHandler(Optional ByRef vStatus As Boolean) As String
    
    'Declare a variable as a FileDialog object.
    Dim fd As String
    Dim vstrPath As Variant 'The user selected path
    
    fd = Application.GetSaveAsFilename(InitialFileName:=Application.ActiveWorkbook.Path & "\" & Strings.Format(Now(), "YYYYMMDD_hhmmss") & ".in", fileFilter:="FedEx Input (*.in), *.in")

    FileHandler = fd
    vstrPath = fd
    
    If vstrPath = False Then
       vStatus = False
    Else
        vStatus = True
    End If
    'Set the object variable to Nothing.
 '   Set fd = Nothing
    
    
End Function


Function NextAWB() As String
Dim AWB As String, SumAWB As String, Current As String
Dim AWBReminder As Double

Current = Modules.Cells(5, 2)
If Strings.Trim(Current) = "" Or Current >= Modules.Cells(4, 2) Then
    NextAWB = ""
    MsgBox ("No tracking# avaliable")
    Exit Function
End If
AWB = Left(Current, 11) + 1
SumAWB = (Left(AWB, 1) * 3)
SumAWB = SumAWB + (Right(Left(AWB, 2), 1) * 1)
SumAWB = SumAWB + (Right(Left(AWB, 3), 1) * 7)
SumAWB = SumAWB + (Right(Left(AWB, 4), 1) * 3)
SumAWB = SumAWB + (Right(Left(AWB, 5), 1) * 1)
SumAWB = SumAWB + (Right(Left(AWB, 6), 1) * 7)
SumAWB = SumAWB + (Right(Left(AWB, 7), 1) * 3)
SumAWB = SumAWB + (Right(Left(AWB, 8), 1) * 1)
SumAWB = SumAWB + (Right(Left(AWB, 9), 1) * 7)
SumAWB = SumAWB + (Right(Left(AWB, 10), 1) * 3)
SumAWB = SumAWB + (Right(Left(AWB, 11), 1) * 1)


    AWBReminder = CStr(CDbl(SumAWB) Mod 11)
    If AWBReminder = 0 Or AWBReminder = 10 Then
        NextAWB = AWB & "0"
    Else
        NextAWB = AWB & AWBReminder
    End If
Modules.Cells(5, 2) = "'" & CStr(NextAWB)
End Function

Function CountryCode(CountryName As String) As String
    With Country
    intRowCount = .UsedRange.Rows.Count
    
    found = ""
    For i = 2 To intRowCount
        If Strings.Trim(Strings.UCase(CountryName)) = Strings.Trim(Strings.UCase(.Cells(i, 2))) Then
            found = .Cells(i, 1)
        End If
    Next i
    
    CountryCode = found
End With

End Function

Sub Tracking_Click()
Dim MasterAWB As String
    With Trans
    intRowCount = .UsedRange.Rows.Count
    intColCount = .UsedRange.Columns.Count

        For intVCounter = 1 To intColCount
            If Trim(CStr(.Cells(3, intVCounter).Value)) = "1123" Then
                    For intCounter = 4 To intRowCount
                        If Trim(CStr(.Cells(intCounter, 1).Value)) <> "" Then
                            .Cells(intCounter, intVCounter).Value = "'" & NextAWB
                        End If
                    Next intCounter
            End If
        Next intVCounter
        End With
End Sub

