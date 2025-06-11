**PURPOSE**:

This VBA code was developed to automate the integration of new expense entries from a ledger file extracted from my ERP system into a specific financial report format. Reporting is carried out monthly, and due to the specific information requirements of my organisation’s reporting template—significantly different from the raw extract provided by the ERP system—manual adaptation was previously required.

The script ensures that only new, non-duplicated entries are added to the report, significantly reducing the risk of manual errors and improving efficiency in periodic financial updates. With the click of a button, it automates an otherwise repetitive and redundant process.

**HOW TO USE**:

Save both files — Financial report.xlsm and Ledger file.xlsx — in the same folder. These files must retain their original names.

Open the file Financial report.xlsm and click the UPDATE REPORT button located in cell B3. Macros must be enabled.

For more details or to review the code logic, open the Visual Basic Editor (Alt + F11) and navigate to the module named FR_Update.

**CODE**:

    Option Explicit
    Sub Financial_report_update()
    Dim Ledger As Worksheet
    Dim FR As Worksheet
    Dim lastLedgerRow As Long
    Dim lastFRRow As Long
    Dim LedgerRange As Range
    Dim FRRange As Range
    Dim newRow As Long
    Dim budgetLine As String
    Dim budgetLineFR As String
    Dim LedgerconcatID As String
    Dim concatCheck As Boolean
    Dim FRConcatID As String
    Dim BLLastRow As Range
    Dim Header As Range
    Dim lastRowInBL As Long
    Dim HeaderBL As Range
    Dim BudgetLHeader1stRow As Long
    Dim BudgetLHeaderlastRow As Long
    
    'Make code faster by disabling screen updating, events and setting calculation to manual to prevent repeated redundant recalculations of formulas in every loop:
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    
    'Opening the data source (ledger); file needs to be named "Ledger file" and saved as xlsx in the same folder as the Financial report file
    Workbooks.Open Filename:=ThisWorkbook.Path & "\" & "Ledger file.xlsx"
    
    ' Assigning the Worksheets we are using as object variables:
    Set Ledger = Workbooks("Ledger file.xlsx").Sheets(1) 'ERP ledger extracts usually generate a single sheet, so I use Sheets(1)
    Set FR = Workbooks("Financial report.xlsm").Sheets("Financial report") 'The sheet in the FR file must be named "Financial report" and saved as xlsm
    
    ' Finding the last row in both sheets and assigning them to variables:
    lastLedgerRow = Ledger.Cells(Ledger.Rows.Count, "A").End(xlUp).Row
    lastFRRow = FR.Cells(FR.Rows.Count, "B").End(xlUp).Row
    
    
    
    
    ' ——— LOOP 1: import data from Ledger ———

    'The purpose of this LOOP is to identify which expenses are in my ledger extract but have not yet been included in my Financial report
    'How? Cross-checking unique ID codes between files - for every transaction recorded in the ERP, a unique journal entry code is created (Entry ID)
    'The Entry ID is present in both the Ledger extract and the Financial report
    'Problem: a single expense may be assigned to multiple budget lines, so the ledger can have multiple rows per Entry ID
    'Solution: concatenate Entry ID + budget line code (which is also included in both files),
    'thus ensuring all rows of expenses are included and under the correct budget line heading in my Financial report
    
    
    ' Loop through each row in the Ledger file
    For Each LedgerRange In Ledger.Range("A2:A" & lastLedgerRow)
    
    ' Create the concatenated ID ("Entry ID" + "budget line code") in the Ledger file
    LedgerconcatID = LedgerRange.Offset(0, 2).Value & "-" & LedgerRange.Offset(0, 6).Value
    
    
    ' Check if the concatenated ID already exists in the Financial report file, that is, if the expense is already included in the report
    concatCheck = False     'Assumes the concatID is not found (False); will be set to True if a match is detected
    For Each FRRange In FR.Range("E2:E" & lastFRRow)
    
        ' Create the concatenated ID in Financial report file
        FRConcatID = FRRange.Value & "-" & FRRange.Offset(0, -3).Value
        
        
    ' If the concatenated ID is found in the Financial report file, means the expense is already in the report and will not be included again
        If FRConcatID = LedgerconcatID Then
            concatCheck = True ' if there is a match, make concatCheck = true
            Exit For 'as soon as a match is found, it stops looking and jumps to the next range; faster code
        End If
    Next FRRange
    
    ' If the concatenated ID is not found in the Financial report file, it's a new expense and must be added
    If Not concatCheck Then 'meaning: "if concatCheck is not True then"
    
        ' Get the budget line for the new expense
        budgetLine = LedgerRange.Offset(0, 6).Value
        
        ' Find the last row of the corresponding budget line in the Financial report file
        Set BLLastRow = FR.Columns("B").Find(What:=budgetLine, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        
        'In case we have a new budget line (not included in the report template yet), the expense will be included in a table at the bottom
        If BLLastRow Is Nothing Then
            
            newRow = lastFRRow + 1
            
        
        'Otherwise, include the expense just after the last added expense for the respective budget line
        Else
            newRow = BLLastRow.Row + 1
                       
        End If
        
        'Insert a new row at the identified position
        
        FR.Rows(newRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
        ' Copy relevant information from Ledger file to Financial report file
        FR.Cells(newRow, "B").Value = LedgerRange.Offset(0, 6).Value    'Budget line
        FR.Cells(newRow, "C").Value = LedgerRange.Offset(0, 1).Value    'Issue date
        FR.Cells(newRow, "D").Value = LedgerRange.Value                 'Entry date
        FR.Cells(newRow, "E").Value = LedgerRange.Offset(0, 2).Value    'Entry ID
        FR.Cells(newRow, "F").Value = LedgerRange.Offset(0, 11).Value   'Invoice number
        FR.Cells(newRow, "G").Value = LedgerRange.Offset(0, 10).Value   'Description
        FR.Cells(newRow, "N").Value = LedgerRange.Offset(0, 12).Value   '#units
        FR.Cells(newRow, "O").Value = LedgerRange.Offset(0, 13).Value   'Unit cost
        FR.Cells(newRow, "P").Value = LedgerRange.Offset(0, 14).Value   'Total cost
        FR.Cells(newRow, "S").Value = LedgerRange.Offset(0, 18).Value   'Activity code
        FR.Cells(newRow, "T").Value = LedgerRange.Offset(0, 20).Value   'FR number
        
        
        'I want the font colour of the new lines to be blue and not Bold
        With FR.Rows(newRow).Font
            .Color = RGB(0, 0, 255)
            .Bold = False
        End With
        
        'I also want to be able to easily spot the newly inserted lines, for later review, thus painting them in light green
        FR.Rows(newRow).Interior.Color = RGB(226, 239, 218)
                    
        
        ' Updating the last Financial report row after the insertion of a new expense before restarting the LOOP
        lastFRRow = FR.Cells(FR.Rows.Count, "B").End(xlUp).Row
    End If
    Next LedgerRange
    
    
    
      ' ——— END of LOOP 1 ———



      Workbooks("Ledger file.xlsx").Close SaveChanges:=False ' Closing ledger file


        
      ' ——— LOOP 2: update SUM formulas for each budget line heading ———

    'The purpose of this LOOP is to update formulas in the Financial report file accordingly to the newly inserted rows in LOOP 1, namely:
    'SUMs of the units (# of units formula) and expenditures amount (Budget executed formula) in each budget line heading
    

        For Each FRRange In FR.Range("P2:P" & lastFRRow)
        
        'The budget line code is under Column B:
        budgetLineFR = FRRange.Offset(0, -14).Value
        
        'Setting the header as the first correspondence found in column B when searching for the budget line:
        Set Header = FR.Columns("B").Find(What:=budgetLineFR, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
        
        BudgetLHeader1stRow = Header.Row 'returns the Row number for the Header
        BudgetLHeaderlastRow = Header.End(xlDown).Row 'returns the Row number for the last expense added under the budget line
        
        'The headings' rows are identified under column A by "H3" (in Font colour White)
        'Adjusting formula update in case there is no expenses under the respective heading - to not include the next heading in the sum formula
        If FR.Cells(BudgetLHeader1stRow, "A").Value = "H3" And IsEmpty(Header.Offset(1, 14).Value) Then
                     
        FR.Cells(BudgetLHeader1stRow, "P").Formula = "=SUM(P" & (BudgetLHeader1stRow + 1) & ":P" & (BudgetLHeaderlastRow - 1) & ")" 'Budget executed formula
        FR.Cells(BudgetLHeader1stRow, "N").Formula = "=SUM(N" & (BudgetLHeader1stRow + 1) & ":N" & (BudgetLHeaderlastRow - 1) & ")" '# of units formula

        ElseIf FR.Cells(BudgetLHeader1stRow, "A").Value = "H3" Then
                     
        FR.Cells(BudgetLHeader1stRow, "P").Formula = "=SUM(P" & (BudgetLHeader1stRow + 1) & ":P" & BudgetLHeaderlastRow & ")" 'Budget executed formula
        FR.Cells(BudgetLHeader1stRow, "N").Formula = "=SUM(N" & (BudgetLHeader1stRow + 1) & ":N" & BudgetLHeaderlastRow & ")" '# of units formula
        
        Else
        
        'do nothing
                      
        End If
                
        
        Next FRRange
        
        
      ' ——— END of LOOP 2 ———



      'Settings back to normal:
      Application.ScreenUpdating = True
      Application.EnableEvents = True
      Application.Calculation = xlCalculationAutomatic



      End Sub



