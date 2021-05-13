Sub Validation_Script()
'
' Validation_Script Macro
' Make sure all the CIF entries are there and they're not obviously wrong.
'
' Keyboard Shortcut: Ctrl+Shift+V
'''''''''''''''''''''''''''''''''''''''''''''''''''
'If there is already and Errors sheet, delete it
   For Each Sheet In ActiveWorkbook.Worksheets
    If Sheet.Name = "Errors" Then
        Application.DisplayAlerts = False
        Worksheets("Errors").Delete
        Application.DisplayAlerts = True
    End If
Next Sheet
'''''''''''''''''''''''''''''''''''''''''''''
    Sheets.Add After:=ActiveSheet 'Add a new sheet for errors
    ActiveSheet.Name = "Errors"
    Range("A1").Value = "CIF ERRORS"
'''''''''''''''''''''''''''''''''''''''''''''''''''
    'Verifies the formula for "Contractor Share of Costs"
    Sheets("Form").Select 'make sure the CIF is selected
    Contractor_Share_Bottom = Range("J68")
    Contractor_Share_Middle = Range("C26")
    If Contractor_Share_Bottom <> Contractor_Share_Middle Then
        With Sheets("Errors")
            .Select
            FindLastCell ("Error Code 01:  Check your formula in C26.  It should reference J68 and the values should match.")
        End With
        Sheets("Form").Select
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Verifies the formula for "TWDB Share of Costs"
    Sheets("Form").Select 'make sure the CIF is selected
    TWDB_Share_Bottom = Range("J51")
    TWDB_Share_Middle = Range("C27")
    If TWDB_Share_Bottom <> TWDB_Share_Middle Then
        With Sheets("Errors")
            .Select
            FindLastCell ("Error Code 02:  Check your formula in C27.  It should reference J51 and the values should match.")
        End With
        Sheets("Form").Select
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''
    'Verifies the formula for "Receivable Share of Costs"
    Sheets("Form").Select 'make sure the CIF is selected
    Receievable_Share_Bottom = Range("J67")
    Receievable_Share_Middle = Range("C28")
    If Receievable_Share_Bottom <> Receievable_Share_Middle Then
        With Sheets("Errors")
            .Select
            FindLastCell ("Error Code 03:  Check your formula in C28.  It should reference J67 and the values should match.")
        End With
        Sheets("Form").Select
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''
    'Verifies the formula for "Total Contract Costs"
    Sheets("Form").Select 'make sure the CIF is selected
    Total_Contract_Costs_Bottom = Range("J69")
    Total_Contract_Costs_Middle = Range("C29")
    If Total_Contract_Costs_Bottom <> Total_Contract_Costs_Middle Then
        With Sheets("Errors")
            .Select
            FindLastCell ("Error Code 04:  Check your formula in C29.  It should reference J69 and the values should match.")
        End With
        Sheets("Form").Select
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''
    'Verifies that there is a contract number unless its new and that the number, if any, is 10 digits
    If IsEmpty(Range("B9").Value) And IsEmpty(Range("I3").Value) Then
        Sheets("Errors").Select
        FindLastCell ("Error Code 05:  No contract number and it's not a new contract.  Compare cell I9 and B9.")
    ElseIf Len(Range("B9")) <> 10 Then
        Sheets("Errors").Select
        FindLastCell ("Error Code 06:  Contract Number should be 10 digits.")
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Verifies that there is a yes or no in the Grant cell
    If LCase(Range("E9").Value) <> "yes" And LCase(Range("E9").Value) <> "no" Then
        FindLastCell ("Error Code 07:  Check Cell E9.  It should say Yes or No.  ")
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''
 'Verifies that it's either payable or receivable.  It can't be both
    If IsEmpty(Range("B10").Value) = "False" And IsEmpty(Range("E10:F10").Value) = "False" Then
        FindLastCell ("Error Code 08:  Compare B10 and E10.  Cannot be both Payable and Receivable.  ")
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Verifies that there is something in the Board Approval Date Field
    Sheets("Form").Select 'make sure the CIF is selected
    If IsEmpty(Range("A13").Value) Then
        FindLastCell ("Error Code 09:  Board Approval Date, cell A13, cannot be blank.  NA if no date.  ")
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Warns if the start date is in the past
    If CDate(Range("C13").Value) <= Now() Then
        FindLastCell ("Warn Code 10:  Start date is in the past.  Cell C13.  ")
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Verifies that the Expiraiton date is not before the Start Date
    If IsEmpty(Range("E13").Value) Then
        FindLastCell ("Error Code 11:  Expiration date cannot be blank.  Cell E13.  ")
    ElseIf CDate(Range("E13").Value) < CDate(Range("C13").Value) Then
        FindLastCell ("Error Code 12:  Expiration date is before start date.  Cell E13.  ")
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Verifes that Most Recent Amendment Execution Date is there if and only if this is an Amendment other than Amendment 1
    If Range("I5") >= 2 And IsEmpty(Range("C14").Value) Then
       FindLastCell ("Error Code 13:  If there has been a previous amendment, put its Execution Date in A14.  Reference I5, which shows this is not the first amendment.")
    End If
    If Range("I5") < 2 And IsEmpty(Range("C14").Value) = False Then
        FindLastCell ("Error Code 14:  If there has been no previous amendment, then cell A14 should be blank.  Reference I5, which shows this is not an Amendment other than Amd 1.  ")
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Original Contract Expiration should not be blank if this is an amendment.  Must be blank if it's not.
    If IsEmpty(Range("I4").Value) And IsEmpty(Range("F14").Value) = False Then
           FindLastCell ("Error Code 15:  Cell I3 says this is not an amendment, so there shouldn't be anything in Original Contract Expiration Date F14.")
    End If
    
    If IsEmpty(Range("I4").Value) = False And IsEmpty(Range("F14").Value) Then
           FindLastCell ("Error Code 16:  Cell I3 says this is an amendment, so there should be an Original Contract Expiration Date ini F14.")
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Proposal number should be 8 digits, four of which are zeros
    If Len(Range("C16")) <> 8 Or Range("C16").Value > 9999 Then
        FindLastCell ("Warn Code 17:  Is this proposal number correct? It either is not 8 digits or the first 4 are not zeros.")
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Vendor ID should be 9 or 11 digits
    If Len(Range("C17")) <> 9 And Len(Range("C17")) <> 11 Then
        FindLastCell ("Warn Code 18:  Make sure this vendor ID number is correct.")
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Other vendor info just should not be blank
    If IsEmpty(Range("C18")) Then 'Vendor Name
        FindLastCell ("Error Code 19:  Vendor name cannot be blank - C18.")
    End If
    '''
    If IsEmpty(Range("C19")) Then 'Vendor Street Address
        FindLastCell ("Error Code 20:  Vendor Street Address cannot be blank - C19.")
    End If
    '''
    If IsEmpty(Range("C20")) Then 'Vendor City/State/Zip
        FindLastCell ("Error Code 21:  Vendor City/State/Zip cannot be blank - C20.")
    End If
    '''
    If IsEmpty(Range("C21")) Then 'Vendor Phone
        FindLastCell ("Warn Code 22:  Vendor phone number probably should not be blank - C21.")
    End If
    '''
    If IsEmpty(Range("C22")) Then 'Vendor Phone
        FindLastCell ("Error Code 23:  Vendor Contract Mgr/Email Address cannot be blank - C22.")
    End If
    '''
    If IsEmpty(Range("C23")) Then 'Vendor Phone
        FindLastCell ("Error Code 24:  Vendor signatory cannot be blank - C23.")
    End If
    
    'This should be the last thing
    Format()
    
End Sub

'From https://techcommunity.microsoft.com/t5/excel/using-vba-to-select-first-empty-row/m-p/28181
Sub FindLastCell(message As String)

Dim LastCell As Range
Dim LastCellColRef As Long

LastCellColRef = 1 'column number to look in when finding last cell

    Set LastCell = Sheets("Errors").Cells(Rows.Count, LastCellColRef).End(xlUp).Offset(1, 0)
    
    'MsgBox LastCell.Address
    LastCell.Value = message
        
Set LastCell = Nothing
Sheets("Form").Select 'make sure the CIF is selected
    
End Sub

'Formats the Errors Sheet to differentiate between warnings and errors
Sub Format()
  Sheets("Errors").Select 'make sure the Errors Sheet is selected
  Columns("A:A").ColumnWidth = 127.86
    Columns("A:A").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Error Code", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Columns("A:A").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Warn Code", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("A1").Select
    Selection.Font.Bold = True
End Sub

