Sub WSH_TRANSACTION_RAW()
'
' WSH_TRANSACTION_RAW Macro
' STEP1: Replace =+ to +
' STEP2: ORD value excluded empty
' STEP3: Paid Search Campaign name is not empty
' STEP4: Revenue lower than $10
' STEP5: Campaign Name Excluded ECRM
' STEP6: Open Generator File
' STEP7: Paste Data
' STEP8: Fill all the columns
' STEP9: Remove duplicate
'   Calculation change to manual
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
'   Name Variables
    Dim WshTrxRPath As String
    Dim WshTrxRFile As String
    Dim LatestTrxRFile As String
    Dim LatestTrxRDate As Date
    Dim WshTrxPaF As Date
'   Specify the path to the folder
    WshTrxRPath = "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\Raw_Data\Transactions\Transactions_DFA\2017"
'   Make sure that the path ends in a backslash
    If Right(WshTrxRPath, 1) <> "\" Then WshTrxRPath = WshTrxRPath & "\"
'   Get the first Excel file from the folder
    WshTrxRFile = Dir(WshTrxRPath & "*.csv", vbNormal)
'   Loop through each Excel file in the folder
    Do While Len(WshTrxRFile) > 0
'   Assign the date/time of the current file to a variable
    WshTrxPaF = FileDateTime(WshTrxRPath & WshTrxRFile)
'   If the date/time of the current file is greater than the latest, recorded date, assign its filename and date/time to variables
    If WshTrxPaF > LatestTrxRDate Then
    LatestTrxRFile = WshTrxRFile
    LatestTrxRDate = WshTrxPaF
    End If
'   Get the next Excel file from the folder
    WshTrxRFile = Dir
    Loop
'   Open the latest file
    Workbooks.Open WshTrxRPath & LatestTrxRFile
'   Select the first cell of the data
    Range("$A12").Select
'   Set up filter
    Selection.AutoFilter
'   Find last row of the data
    Dim LastRow As Long
    LastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row
'   Replace "=+" to "+"
    Cells.Replace What:="=+", Replacement:="+", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'   Filter out null from "ORD Value"
    ActiveSheet.Range("A12:AH" & LastRow).AutoFilter Field:=10, Criteria1:= _
        "<>*Null*", Operator:=xlAnd, Criteria2:="<>*--*"
'   Filter out ECRM
    ActiveSheet.Range("A12:AH" & LastRow).AutoFilter Field:=5, Criteria1:= _
        "<>*ECRM*", Operator:=xlAnd
'   Filter out Revenue <= $10
    ActiveSheet.Range("A12:AH" & LastRow).AutoFilter Field:=11, Criteria1:=">10", _
        Operator:=xlAnd
'   Select data range
    Range("A13:AH" & LastRow).Select
    Selection.Copy
End Sub

Sub WSH_TRANSACTION()
'   Name all the variables for transaction file
    Dim WshTrxFile As String
    Dim WshTrxPath  As String
    Dim LatestTrxFile As String
    Dim LatestTrxDate As Date
    Dim WshTrx As Date
    WshTrxPath = "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\Transactions\2017"
    If Right(WshTrxPath, 1) <> "\" Then WshTrxPath = WshTrxPath & "\"
    WshTrxFile = Dir(WshTrxPath & "*.xlsb", vbNormal)
'   Loop through each Excel file in the folder
    Do While Len(WshTrxFile) > 0
'   Assign the date/time of the current file to a variable
    WshTrx = FileDateTime(WshTrxPath & WshTrxFile)
'   If the date/time of the current file is greater than the latest, recorded date, assign its filename and date/time to variables
    If WshTrx > LatestTrxDate Then
    LatestTrxFile = WshTrxFile
    LatestTrxDate = WshTrx
    End If
'   Get the next Excel file from the folder
    WshTrxFile = Dir
    Loop
'   Name the transaction file
    Dim NewTransactionFile As String
    NewTransactionFile = WshTrxPath & "WSH_" & "04012016" & "-" & Format(Date - WeekDay(Date), "mmddyyyy") & "_" & "Transactions" & ".xlsb"
'   Copy last week transaction file and rename to current week
    CreateObject("Scripting.FileSystemObject").CopyFile WshTrxPath & LatestTrxFile, NewTransactionFile
'   Open Current Week Transaction File
    Workbooks.Open NewTransactionFile
'   Active Transaction tab
    Worksheets("Transactions").Activate
'   Go to the last empty cell of raw data first column
    Range("$A2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("F1").Select 'Try ActiveCell.Offset(1, 5).Select
    Worksheets("Transactions").Paste
'   Fill the data on the left
'   Range("$A2:$E2").Select
'   Selection.Copy
'   Range("$F2").Select
'   Selection.End(xlDown).End(xlToLeft).Select
'   Range(Selection, Selection.End(xlUp)).Select
'   ActiveSheet.Paste
'   Fill the data on the right
'   Range("$AN2:$CA2").Select
'   Selection.Copy
'   Range("$AM2").Select
'   Selection.End(xlDown).Select
'   ActiveCell.Offset(0, 2).Select 'Try was Range(B1)
'   Range(Selection, Selection.End(xlUp)).Select
'   ActiveSheet.Paste
'   Duplicate Remove  (Remove the new dates)
'   ActiveSheet.Range("$A$1:$CA$60000").RemoveDuplicates Columns:=15, Header:=xlYes
'   Calculate the file
'   Application.Calculation = xlAutomatic
'   Application.ScreenUpdating = True
'   Find #N/A
'   Cells.find(What:="#N/A", After:=ActiveCell, LookIn:=xlValues, LookAt:= _
'   xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
'   Workbooks(LatestTrxRFile).Close Savechanges:=False
'   ActiveWorkbook.Save
End Sub
Sub WSH_CAMPAIGN_WK()
'
'   STEP1: Clean up display data(include cost adjustment), paste, fill the formulas
'   STEP2: Copy SEM Data, paste, fill the formulas
'   STEP3: Paste transactions data, fill the formulas
'   Declare the varibles
    Dim WshDisplayRPath As String
    Dim WshDisplayRFile As String
    Dim LatestDisplayRFile As String
    Dim LatestDisplayRDate As Date
    Dim WshDisplay As Date
'   Specify the path to the folder
    WshDisplayRPath = "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\Raw_Data\Display\Display_WK\2017"
'   Make sure that the path ends in a backslash
    If Right(WshDisplayRPath, 1) <> "\" Then WshDisplayRPath = WshDisplayRPath & "\"
'   Get the first Excel file from the folder
    WshDisplayRFile = Dir(WshDisplayRPath & "*.csv", vbNormal)
'   Loop through each Excel file in the folder
    Do While Len(WshDisplayRFile) > 0
'   Assign the date/time of the current file to a variable
    WshDisplay = FileDateTime(WshDisplayRPath & WshDisplayRFile)
'   If the date/time of the current file is greater than the latest, recorded date, assign its filename and date/time to variables
    If WshDisplay > LatestDisplayRDate Then
    LatestDisplayRFile = WshDisplayRFile
    LatestDisplayRDate = WshDisplay
    End If
'   Get the next Excel file from the folder
    WshDisplayRFile = Dir
    Loop
'   Open the latest file
    Workbooks.Open WshDisplayRPath & LatestDisplayRFile
'   Top 10 rows delete
    Rows("1:11").Delete
    Range("$A1").Select
'   Apply AutoFilter
    If ActiveSheet.AutoFilterMode Then Selection.AutoFilter
    ActiveCell.CurrentRegion.Select
'   Delete Filter data without header
    With Selection
        .AutoFilter
        .AutoFilter Field:=5, Criteria1:= _
        "=*not set*", Operator:=xlOr, Criteria2:="=*--*"
        .Offset(1, 0).Select
    End With
    Dim numRows As Long, numColumns As Long
    numRows = Selection.Rows.Count
    numColumns = Selection.Columns.Count
    Selection.Resize(numRows - 1, numColumns).Select
    With Selection
    .SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End With
'   Clear All the filters
    Cells.AutoFilter
'   Find last row of the data
    Dim DisplayLastRow As Long
    DisplayLastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row
'   Filter out Media IQ
'   ActiveSheet.Range("A1:R" & DisplayLastRow).AutoFilter Field:=3, Criteria1:="=*MEDIA IQ*" _
'       , Operator:=xlAnd
'   Zero out all the organial data (Test)
'   Range("$J2:$J" & DisplayLastRow).FormulaR1C1 = "0"
'   Filter out GDN
    ActiveSheet.Range("$A1:$R" & DisplayLastRow).AutoFilter Field:=3, Criteria1:="=*GDN*" _
        , Operator:=xlAnd
'   Zero out all the organial data (Test)
    Range("$H2:$J" & DisplayLastRow).FormulaR1C1 = "0"
'   ActiveSheet.Range("A1:R" & DisplayLastRow).AutoFilter Field:=3, Criteria1:="=*MEDIA IQ*" _
       , Operator:=xlAnd
'   Open CostAdjustment Raw Data File
'   Workbooks.Open Filename:= _
'        "\\CLIENTSERVER\analytics\Report_Adjustments\WSH\WSH_Adjustment_New.xlsx"
'   Worksheets("Weekly").Activate
'   Find the last of the cost adjustment file
'   Dim CostLastRow As Long
'   CostLastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row
'   ActiveSheet.Range("$A1:$H" & CostLastRow).AutoFilter Field:=3, Criteria1:="=Media IQ" _
'       , Operator:=xlAnd
'   Copy Media IQ current week cost
'   Range("$G2").End(xlDown).Copy
'   Go back to display raw file
'   Workbooks(LatestDisplayRFile).Activate
'   Range("$J2").End(xlDown).Select
'   ActiveSheet.Paste
'   Clear the filter
    Cells.AutoFilter
'   Filter out ECRM
    ActiveSheet.Range("$A1:$R" & DisplayLastRow).AutoFilter Field:=2, Criteria1:= _
        "<>*ECRM*", Operator:=xlAnd
'   Copy Data
    Range("$A2:$R" & DisplayLastRow).Select
    Selection.Copy
'   Workbooks("WSH_Adjustment_New.xlsx").Close Savechanges:=False
'   Find the latest report generator file
    Dim WshReportGenPath As String
    Dim WshReportGenFile As String
    Dim LatestReportGenFile As String
    Dim LatestReportGenDate As Date
    Dim WshReportGenerator As Date
'   Specify the path to the folder
    WshReportGenPath = "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\ReportGenerator_WK\2017"
'   Make sure that the path ends in a backslash
    If Right(WshReportGenPath, 1) <> "\" Then WshReportGenPath = WshReportGenPath & "\"
'   Get the first Excel file from the folder
    WshReportGenFile = Dir(WshReportGenPath & "*.xlsb", vbNormal)
'   Loop through each Excel file in the folder
    Do While Len(WshReportGenFile) > 0
'   Assign the date/time of the current file to a variable
    WshReportGenerator = FileDateTime(WshReportGenPath & WshReportGenFile)
'   If the date/time of the current file is greater than the latest, recorded date, assign its filename and date/time to variables
    If WshReportGenerator > LatestReportGenDate Then
    LatestReportGenFile = WshReportGenFile
    LatestReportGenDate = WshReportGenerator
    End If
'   Get the next Excel file from the folder
    WshReportGenFile = Dir
    Loop
'   Workbooks.Open PccTrxPath & LatestTrxFile
    Dim NewReportGeneratorFile As String
    NewReportGeneratorFile = WshReportGenPath & "WSH_" & Format(Date - WeekDay(Date) - 6, "mmddyyyy") & "-" & Format(Date - WeekDay(Date), "mmddyyyy") & "_" & "ReportGenerator-NewReportFormat" & ".xlsb"
'   Copy last week transaction file and rename to curent week
    CreateObject("Scripting.FileSystemObject").CopyFile WshReportGenPath & LatestReportGenFile, NewReportGeneratorFile
'   Open Current Week Transaction File
    Workbooks.Open NewReportGeneratorFile
    Worksheets("Campaigns_Wk").Activate
'   Go to the last cell without data
    Range("$E2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
'   Open SEM Raw Data File
'   Declare the varibles
    Dim WshSEMRPath As String
    Dim WshSEMRFile As String
    Dim LatestSEMRFile As String
    Dim LatestSEMRDate As Date
    Dim WshSEM As Date
'   Specify the path to the folder
    WshSEMRPath = "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\Raw_Data\SEM\SEM_WK\2017"
'   Make sure that the path ends in a backslash
    If Right(WshSEMRPath, 1) <> "\" Then WshSEMRPath = WshSEMRPath & "\"
'   Get the first Excel file from the folder
    WshSEMRFile = Dir(WshSEMRPath & "*.csv", vbNormal)
'   Loop through each Excel file in the folder
    Do While Len(WshSEMRFile) > 0
'   Assign the date/time of the current file to a variable
    WshSEM = FileDateTime(WshSEMRPath & WshSEMRFile)
'   If the date/time of the current file is greater than the latest, recorded date, assign its filename and date/time to variables
    If WshSEM > LatestSEMRDate Then
    LatestSEMRFile = WshSEMRFile
    LatestSEMRDate = WshSEM
    End If
'   Get the next Excel file from the folder
    WshSEMRFile = Dir
    Loop
'   Open the latest file
    Workbooks.Open WshSEMRPath & LatestSEMRFile
'   Copy Data
    Dim SEMLastRow As Long
    SEMLastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row
    ActiveSheet.Range("$A12:$R" & SEMLastRow).AutoFilter Field:=1, Criteria1:= _
         "<>*Grand Total*", Operator:=xlAnd
    Range("A13:R" & SEMLastRow).Select
    Selection.Copy
    Workbooks.Open NewReportGeneratorFile
    Worksheets("Campaigns_Wk").Activate
'   Go to the last cell without data
    Range("$E2").Select
    Selection.End(xlDown).Select
'   Range("A1") means "Go to Current column"; Offset(1,0)means "Go to next row"
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
'   Fill the data On the Right
    Range("$W2:$BF2").Select
    Selection.Copy
'   Go to the last cell of one column with blanks
    Dim lngLastRow As Long
    lngLastRow = Columns(22).Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Select
    ActiveCell.Offset(0, 0).Range("B1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
'   Change Date
    Dim LastSunday As String
    LastSunday = Format(Date - WeekDay(Date) - 6, "mm/dd/yyyy")
    Range("$D2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Selection.FormulaR1C1 = LastSunday
    ActiveCell.Copy
    Range("$E2").End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
'   Fill the data on the left
    Range("$A2:$C2").Copy
    Range("$E2").End(xlDown).Offset(0, -4).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
'   Active Transaction tab
    Dim WshTrxFile As String
    Dim WshTrxPath  As String
    Dim LatestTrxFile As String
    Dim LatestTrxDate As Date
    Dim WshTrx As Date
    WshTrxPath = "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\Transactions\2017"
    If Right(WshTrxPath, 1) <> "\" Then WshTrxPath = WshTrxPath & "\"
    WshTrxFile = Dir(WshTrxPath & "*.xlsb", vbNormal)
'   Loop through each Excel file in the folder
    Do While Len(WshTrxFile) > 0
'   Assign the date/time of the current file to a variable
    WshTrx = FileDateTime(WshTrxPath & WshTrxFile)
'   If the date/time of the current file is greater than the latest, recorded date, assign its filename and date/time to variables
    If WshTrx > LatestTrxDate Then
    LatestTrxFile = WshTrxFile
    LatestTrxDate = WshTrx
    End If
'   Get the next Excel file from the folder
    WshTrxFile = Dir
    Loop
    Workbooks.Open WshTrxPath & LatestTrxFile
    Worksheets("Summary").Activate
'   Refresh Pivot Tables on the worksheet
    Dim xTable As PivotTable
    For Each xTable In Application.ActiveSheet.PivotTables
    xTable.RefreshTable
    Next
    Dim TrxLastRow As Long
    TrxLastRow = ActiveSheet.Range("B" & ActiveSheet.Rows.Count).End(xlUp).Row
'   ActiveSheet.Range("A3:Q" & TrxRLastRow).AutoFilter Field:=1, Criteria1:= _
'       "<>*(blank)*", Operator:=xlAnd
    Range("A5:S" & TrxLastRow).Select
    Selection.Copy
    Workbooks.Open NewReportGeneratorFile
'   Active Transaction tab
    Worksheets("Trx_Summary").Activate
    Range("$D2").Select
    ActiveSheet.Paste
'   Fill the data on the left
    Range("$A2:$C2").Select 'Column D,E are blanks
    Selection.Copy
    Dim TrxLastRow2 As Long
    TrxLastRow2 = ActiveSheet.Range("D" & ActiveSheet.Rows.Count).End(xlUp).Row
    Range("A2:C" & TrxLastRow2).Select
    ActiveSheet.Paste
    Workbooks(LatestDisplayRFile).Close Savechanges:=False
    Workbooks(LatestSEMRFile).Close Savechanges:=False
End Sub

Sub WSH_MTD()
'   Open Display MTD Raw Data File
    Workbooks.Open Filename:= _
         "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\Raw_Data\Display\Display_MTD\WSH_12172017_12232017_Display-MTD.csv"
    Worksheets("WSH_12172017_12232017_Display-M").Activate
'   Filter out Media IQ
    ActiveSheet.Range("$A11:$I150").AutoFilter Field:=4, Criteria1:="=*MEDIA IQ*" _
        , Operator:=xlAnd
'   Zero out all the organial data
    Range("$G12:$G150").SpecialCells(xlCellTypeConstants).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
'   Fill the deleted data cells with zeros
    Selection.FormulaR1C1 = "0"
'   Filter out GDN
    ActiveSheet.Range("$A11:$I150").AutoFilter Field:=4, Criteria1:="=*GDN*" _
        , Operator:=xlAnd
'   Zero out all the organial data
    Range("$E12:$G150").SpecialCells(xlCellTypeConstants).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
'   Fill the deleted data cells with zeros
    Selection.FormulaR1C1 = "0"
'   Adjustment apply for Facebook
'   ActiveSheet.Range("$A12:$I100").AutoFilter Field:=4, Criteria1:="=*Facebook*" _
'       , Operator:=xlAnd
'   ActiveSheet.Range("$A12:$I100").AutoFilter Field:=3, Criteria1:="=*Main*" _
'       , Operator:=xlAnd
'   Range("$E12").End(xlDown).Select
'   ActiveCell.Value="7751" 'Impression
'   ActiveCell.Offset(0, 1).Value = "36" 'Clicks
'   ActiveCell.Offset(0, 2).Value = "54.44" 'Cost
'   Clear All the filters
    Cells.AutoFilter
'   Find Values or Variables from certain Columns
'   Const WHAT_TO_FIND As String = "Media IQ"
'   Set FoundCell = ActiveSheet.Range("$D12:$D150").Find(What:=WHAT_TO_FIND)
'   If Not FoundCell Is Nothing Then
'   IF find it, fill the cost adjustment at the first cell
'       FoundCell.Offset(0, 3).Value = "2165" ' Check whether its still working
'   Else
'   If did not find it, send a message says"Not found"
'       MsgBox (WHAT_TO_FIND & "not found")
'   End If
'   Filter out ECRM
    ActiveSheet.Range("$A11:$I150").AutoFilter Field:=3, Criteria1:= _
        "<>*Natural*", Operator:=xlAnd, Criteria2:="<>*DART*"
    ActiveSheet.Range("$A11:$I150").AutoFilter Field:=4, Criteria1:= _
        "<>*worldmedia*", Operator:=xlAnd
    ActiveSheet.Range("$A11:$I150").AutoFilter Field:=1, Criteria1:="<>*Grand Total*" _
        , Operator:=xlAnd
'   Copy Data
    Range("$A12:$I150").SpecialCells(xlCellTypeConstants).Select
    Selection.Copy
'   Open ReportGenerator File
    Dim CurPath As String
    Dim CurReportGen As String
    CurPath = "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\ReportGenerator_WK\2017\"
    CurReportGen = "WSH_12172017-12232017_ReportGenerator-NewReportFormat.xlsb"
    Workbooks.Open CurPath & CurReportGen
'   Active worksheet
    Worksheets("DFA_MTD").Activate
'   Go to the last cell without data
    Range("$E2").End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
'   Open SEM Raw Data File
    Workbooks.Open Filename:= _
        "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\Raw_Data\SEM\SEM_MTD\2017\WSH_12172017-12232017_SEM-MTD.csv"
    Worksheets("WSH_12172017-12232017_SEM-MTD").Activate
    ActiveSheet.Range("$A$11:$I$2000").AutoFilter Field:=1, Criteria1:= _
        "<>*Grand Total*", Operator:=xlAnd
    Range("$A12:$I2000").SpecialCells(xlCellTypeConstants).Select
    Selection.Copy
'   Open ReportGenerator File
    Workbooks.Open CurPath & CurReportGen
'   Active worksheet
    Worksheets("DFA_MTD").Activate
'   Go to the last cell without data
    Range("$E2").Select
    Selection.End(xlDown).Select
'   Offset(1,0)means "Go to next row"
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
'   Fill the data on the left
    Range("$A2:$D2").Select
    Selection.Copy
    Range("$E2").Select
    Selection.End(xlDown).End(xlToLeft).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
'   Fill the data On the Right
    Range("$N2:$P2").Select
    Selection.Copy
'   Go to the last cell of one column with blanks
    Dim lngLastRow As Long
    lngLastRow = Columns(13).Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Select
    ActiveCell.Offset(0, 0).Range("B1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
End Sub
Sub WSH_GA_CE_AT_ECRM()
'   Open Transaction_GA Raw Data File
    Workbooks.Open Filename:= _
         "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\Raw_Data\Transactions\Transactions_GA\2017\WSH_12172017-12232017_Transactions_GA.csv"
    Worksheets("WSH_12172017-12232017_Transacti").Activate
'   Delete the data around
    Range("$A14").End(xlDown).Offset(1).Resize(ActiveSheet.UsedRange.Rows.Count).EntireRow.Delete
'   Copy data
    Range("$A15:$E150").SpecialCells(xlCellTypeConstants).Select
    Selection.Copy
'   Open ReportGenerator File
    Dim CurPath As String
    Dim CurReportGen As String
    CurPath = "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\ReportGenerator_WK\2017\"
    CurReportGen = "WSH_12172017-12232017_ReportGenerator-NewReportFormat.xlsb"
    Workbooks.Open CurPath & CurReportGen
'   Active worksheet
    Worksheets("GA_Campaigns").Activate
    Range("$E2").End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
'   Fill the data on the left
    Range("$A2:$D2").Select
    Selection.Copy
    Range("$E2").End(xlDown).End(xlToLeft).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
'   Fill the data On the Right
    Range("$J2:$P2").Select
    Selection.Copy
'   Go to the last cell of one column with blanks
    Dim lngLastRow As Long
    lngLastRow = Columns(9).Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Select
    ActiveCell.Offset(0, 0).Range("B1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
'   Open Transactions-Cross_Env Raw Data File
    Workbooks.Open Filename:= _
         "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\Raw_Data\Transactions\Transactions_Cross-Env\2017\WSH_12172017-12232017_Transactions-Cross_Env.csv"
    Worksheets("WSH_12172017-12232017_Transacti").Activate
'   Filter out grand total
    ActiveSheet.Range("$A$12:$T$250").AutoFilter Field:=1, Criteria1:= _
         "<>*Grand Total*", Operator:=xlAnd
'   Copy Data
    Range("$A13:$T250").SpecialCells(xlCellTypeConstants).Select
    Selection.Copy
'   Open ReportGenerator File
    Workbooks.Open CurPath & CurReportGen
'   Active worksheet
    Worksheets("Cross-Env").Activate
'   Go to the last cell without data
    Range("$F2").End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
'   Fill the data on the left
    Range("$A2:$E2").Select
    Selection.Copy
    Range("$F2").End(xlDown).End(xlToLeft).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
'   Fill the data On the Right
    Range("$Z2:$AH2").Select
    Selection.Copy
    Range("$Y2").End(xlDown).Select
    ActiveCell.Offset(0, 0).Range("B1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
'   Open Attribution Raw Data File
    Workbooks.Open Filename:= _
       "\\CLIENTSERVER\analytics\WSH\Reporting\Perf_Trx\Report_Generator\Raw_Data\Attribution\Weekly\WSH_12172017-12232017_Attribution.csv"
    Worksheets("WSH_12172017-12232017_Attributi").Activate
'   Delete the data after the main data
    Range("$A7").End(xlDown).Offset(1).Resize(ActiveSheet.UsedRange.Rows.Count).EntireRow.Delete
    ActiveSheet.Range("$A$8:$D$20").SpecialCells(xlCellTypeConstants).Select 'Check whether fix the paste up problem
    Selection.Copy
'   Open ReportGenerator File
    Workbooks.Open CurPath & CurReportGen
    Worksheets("Model_Wk").Activate
'   Go to the last cell without data
    Range("$B2").End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
'   Last Saturday
    Dim LastSat As String
    LastSat = Format(Date - WeekDay(Date), "mm/dd/yyyy")
'   Go to the last cell without data
    Range("$A2").End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Selection.FormulaR1C1 = LastSat
    ActiveCell.Copy
    Range("$B2").End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
'   Fill the data On the Right
    Range("$F2").Select
    Selection.End(xlDown).Copy
    Range("$E2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
'   Open ECRM Raw Data File
    Workbooks.Open Filename:= _
         "\\CLIENTSERVER\analytics\Report_Adjustments\ECRM\Rawdata_File.xlsx"
    Worksheets("WSH_WK").Activate
'   Copy Data
    Range("$L2").Select
'   Go to the last cell of one column with blanks
    Dim lngLastRow1 As Long
    lngLastRow1 = Columns(12).Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Select
    ActiveCell.Offset(0, 0).Range("A1").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.Copy
'   Open ReportGenerator File
    Workbooks.Open CurPath & CurReportGen
    Worksheets("ECRM_Wk").Activate
'   Go to the last cell without data
    Range("$A2").End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
End Sub

