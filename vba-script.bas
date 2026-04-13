' =============================================================================
' PROJECT: Payment Order Reconciliation Automation
' DESCRIPTION:
' Automates reconciliation of payment order reports across multiple branches.
' Pulls data from GL and PO reports, extracts key balances and totals,
' and consolidates results into summary sheets.
' =============================================================================

Sub PO()

    ' =========================
    ' Variable Declarations
    ' =========================
    
    Dim InSourcePath As String
    Dim InSourceWorkBook As Workbook
    Dim SourcePath As String
    Dim SourceWorkBook As Workbook
    Dim Filename As String
    Dim polocation As String
    Dim PoCode As String
    Dim PoClosing As String
    Dim MyWorkbook As Workbook
    Dim InCode As String
    Dim InValue As String
    
    Dim TotalREVA As Double
    Dim TotalINIT As Double

    ' -------------------------
    ' Instrument Totals
    ' -------------------------
    
    ' BCA - BANKER'S CHEQUE - ACCOUNT
    Dim TotalLocationBCA As String, TotalAmtBCA As String, TotalValueBCA As Double
    
    ' BCG - BANKER'S CHEQUE - GL
    Dim TotalLocationBCG As String, TotalAmtBCG As String, TotalValueBCG As Double
    
    ' BCC - BANKER'S CHEQUE - CHEQUE
    Dim TotalLocationBCC As String, TotalAmtBCC As String, TotalValueBCC As Double
    
    ' BCW - BANKER'S CHEQUE - WALKIN
    Dim TotalLocationBCW As String, TotalAmtBCW As String, TotalValueBCW As Double
    
    ' TTW - TELEGRAPHIC TRANSFER - WALKIN
    Dim TotalLocationTTW As String, TotalAmtTTW As String, TotalValueTTW As Double
    
    ' TTA - TELEGRAPHIC TRANSFER - ACCOUNT
    Dim TotalLocationTTA As String, TotalAmtTTA As String, TotalValueTTA As Double
    
    ' TTG - TELEGRAPHIC TRANSFER - GL
    Dim TotalLocationTTG As String, TotalAmtTTG As String, TotalValueTTG As Double
    
    ' EAG - EXPRESS AGAINST GL
    Dim TotalLocationEAG As String, TotalAmtEAG As String, TotalValueEAG As Double

    ' -------------------------
    ' Control Variables
    ' -------------------------
    
    Dim numbers(1 To 183) As String
    Dim x As Integer
    Dim y As Variant
    Dim fileDate As String
    Dim startingRow As Integer
    Dim startingRowIn As Integer

    ' -------------------------
    ' Find Ranges
    ' -------------------------
    
    Dim foundRangePO As Range
    Dim foundRangeInLand As Range
    
    Dim foundRangeBCA As Range
    Dim foundRangeBCG As Range
    Dim foundRangeBCC As Range
    Dim foundRangeBCW As Range
    Dim foundRangeTTW As Range
    Dim foundRangeTTA As Range
    Dim foundRangeTTG As Range
    Dim foundRangeEAG As Range

    ' =============================================================================
    ' USER INPUT
    ' =============================================================================
    
    fileDate = InputBox("Enter the ending date for the files(eg.,14-Feb-24):", "File Ending Date")
    
    If fileDate = "" Then
        MsgBox "Program Cancelled.", vbInformation
        Exit Sub
    End If

    MsgBox "Please wait, Program is Running...This will take 2-3 mins", vbInformation, "Processing"

    startingRow = 3
    startingRowIn = 2

    ' =============================================================================
    ' BRANCH LIST (STATIC ARRAY)
    ' =============================================================================
    
    numbers(1) = "001"
    numbers(2) = "002"
    numbers(3) = "003"
    ' ... (collapsed)
    numbers(183) = 183

    ' =============================================================================
    ' MAIN LOOP
    ' =============================================================================
    
    For x = 1 To 183

        y = numbers(x)

        ' =============================================================================
        ' GL FILE PROCESSING
        ' =============================================================================
        
        SourcePath = "<LOCAL_PATH>\GL\" & y & "_GL_REPORT_" & fileDate & ".xlsx"

        Application.ScreenUpdating = False

        If Dir(SourcePath, vbDirectory) <> "" Then

            Set SourceWorkBook = Workbooks.Open(SourcePath, ReadOnly:=True)

            SourceWorkBook.Sheets(1).Select
            Selection.Range("A1").Select
            Selection.CurrentRegion.Cut

            Workbooks("PO_CHECKER.xlsm").Activate
            ThisWorkbook.Worksheets(3).Select
            ActiveSheet.Range("A1").Select
            ActiveSheet.Paste

            Selection.Columns.AutoFit
            Range("A1:I7").Delete Shift:=xlUp
            Range("B10").Select

            SourceWorkBook.Close savechanges:=False

            ' -------------------------
            ' FIND PO & INLAND VALUES IN THE GL REPORT
            ' -------------------------
            
            Set foundRangePO = Cells.Find("POCODEXXXX", LookIn:=xlFormulas)
            
            If Not foundRangePO Is Nothing Then
                PoCode = foundRangePO.Address
                PoClosing = Range(PoCode).Offset(0, 7).Value
            Else
                PoClosing = "null"
            End If

            Set foundRangeInLand = Cells.Find("INWARDCODEXXX", LookIn:=xlFormulas2)
            
            If Not foundRangeInLand Is Nothing Then
                InCode = foundRangeInLand.Address
                InValue = Range(InCode).Offset(0, 7).Value
            Else
                InValue = "null"
            End If

            Cells.Clear

            Worksheets("RESULTS").Cells(startingRow + x, 1).Value = y
            Worksheets("RESULTS").Cells(startingRow + x, 3) = PoClosing
            Worksheets("RESULTS").Cells(startingRow + x, 4) = InValue

        Else

            Worksheets("RESULTS").Cells(startingRow + x, 1).Value = y
            Worksheets("RESULTS").Cells(startingRow + x, 3) = "GL_Report_Not_Found"
            Worksheets("RESULTS").Cells(startingRow + x, 4) = "Investigate"

        End If

        ' =============================================================================
        ' PAYMENT ORDER FILE PROCESSING
        ' =============================================================================
        
        InSourcePath = "<LOCAL_PATH>\PO\" & y & "_POREPORT_" & fileDate & ".xlsx"

        If Dir(InSourcePath, vbDirectory) <> "" Then

            Set InSourceWorkBook = Workbooks.Open(InSourcePath, ReadOnly:=True)

            InSourceWorkBook.Sheets(1).Select
            Range("A11").Select
            Selection.CurrentRegion.Cut

            Workbooks("PO_CHECKER.xlsm").Activate
            ThisWorkbook.Worksheets(3).Select
            ActiveSheet.Range("A1").Paste

            Selection.Columns.AutoFit

            InSourceWorkBook.Close savechanges:=False

            ' =============================================================================
            ' INSTRUMENT EXTRACTION (ALL TYPES)
            ' =============================================================================
            ' NOTE: Repeated the logic below for all other instruments

            ' Example (BCA)
            Set foundRangeBCA = Cells.Find("Instrument Type: BCA - BANKER'S CHEQUE - ACCOUNT")

            If Not foundRangeBCA Is Nothing Then
                TotalLocationBCA = foundRangeBCA.Address
                Range(TotalLocationBCA).Offset(2, 0).Select
                Selection.CurrentRegion.Find("Total", LookAt:=xlWhole).Activate
                TotalAmtBCA = ActiveCell.Address
                TotalValueBCA = Range(TotalAmtBCA).Offset(0, 1).Value
            Else
                TotalValueBCA = 0
            End If

            ' -------------------------
            ' CALCULATIONS
            ' -------------------------
            
            Range("P70").Value = "=SUMIF(H:H,""INIT"",E:E)"
            TotalINIT = Range("P70").Value

            Range("P71").Value = "=SUMIF(H:H,""REVA"",E:E)"
            TotalREVA = Range("P71").Value

            ' -------------------------
            ' OUTPUT RESULTS
            ' -------------------------
            
            Cells.Clear

            Worksheets("RESULTS").Cells(startingRow + x, 8) = TotalValueBCA
            Worksheets("RESULTS").Cells(startingRow + x, 9) = TotalValueBCG
            Worksheets("RESULTS").Cells(startingRow + x, 10) = TotalValueBCC
            Worksheets("RESULTS").Cells(startingRow + x, 11) = TotalValueBCW
            Worksheets("RESULTS").Cells(startingRow + x, 12) = TotalValueTTW
            Worksheets("RESULTS").Cells(startingRow + x, 13) = TotalValueTTA
            Worksheets("RESULTS").Cells(startingRow + x, 14) = TotalValueTTG
            Worksheets("RESULTS").Cells(startingRow + x, 15) = TotalValueEAG

            Worksheets("ZOOM_IN").Cells(startingRowIn + x, 1) = y
            Worksheets("ZOOM_IN").Cells(startingRowIn + x, 3) = TotalINIT
            Worksheets("ZOOM_IN").Cells(startingRowIn + x, 4) = TotalREVA

        Else

            Worksheets("RESULTS").Cells(startingRow + x, 8) = "PO Report Not Found"
            Worksheets("RESULTS").Cells(startingRow + x, 13) = "Investigate"
            Worksheets("RESULTS").Cells(startingRow + x, 1).Value = y

            Worksheets("ZOOM_IN").Cells(startingRowIn + x, 1) = y

        End If

        Worksheets("RESULTS").Cells(1, 2).Value = fileDate

    Next x

    MsgBox "Program has Finished Running.", vbInformation, "Finished"

End Sub
