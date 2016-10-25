Option Explicit

' Comment the first declaration and uncomment the second for synchronous request class to be used
'Dim bbControlStatic As New DataControl_Static

Dim bbControlSync As New DataControl_Sync
Dim bbControlSync_Y1 As New DataControl_Sync
Dim g_stopDate As Date
'Dim bbControlHistoric As New DataControl_Hist
'

Public Function GetFALScores(rTickers As Range, lYear As Long, sPeriod As String, sConsolidated As String, Optional bHorizontal As Boolean, Optional bExcludeTits As Boolean) As Variant
'   Range formula on the s/sheet
Dim ii As Long, jj As Long, lNumSecs As Long, lNumLevelCols As Long, lNumMetricsCols As Long, lResRow As Long
Dim sSecurity() As String, sOut() As String, bOK As Boolean, lRowOff As Integer
Dim vSecurity As Variant, sYear As Variant, vLevels As Variant, vMetrics As Variant, vTitles As Variant, vMetricsTitles As Variant, vOut() As Variant
Dim dicTickers As Dictionary


On Error GoTo shagged

    g_stopDate = #11/30/2018#    'Adjust to suit
    If Date >= g_stopDate Then
       MsgBox "This code cannot be executed after " & g_stopDate & ".  Please contact GMT Research."
       Exit Function
    End If
    
    Application.StatusBar = "Getting data from Bloomberg..."
    If bHorizontal Then
        lNumSecs = rTickers.Rows.Count
    Else
        lNumSecs = rTickers.Columns.Count
    End If
    ReDim sSecurity(0 To lNumSecs - 1)
    For ii = 0 To lNumSecs - 1
        If bHorizontal Then
            sSecurity(ii) = UCase(rTickers.Cells(ii + 1, 1).Value) & " Equity"
        Else
            sSecurity(ii) = UCase(rTickers.Cells(1, ii + 1).Value) & " Equity"
        End If
    Next ii
    vSecurity = sSecurity
    sYear = CStr(lYear)
   
    bOK = ReturnFALResults(vSecurity, sYear, sPeriod, sConsolidated, vLevels, vMetrics, vTitles, vMetricsTitles)

    Set dicTickers = New Dictionary
    For ii = 0 To UBound(vSecurity) '   Create dictionary of the TICKERS
        If (Not IsEmpty(vLevels(ii, 16))) Then
            If (Not dicTickers.Exists(vLevels(ii, 16))) Then
                dicTickers.Add vLevels(ii, 16), ii
            End If
        End If
    Next ii

    lNumLevelCols = UBound(vTitles)
    lNumMetricsCols = UBound(vMetricsTitles)
    
    ReDim vOut(0 To lNumSecs, 0 To lNumLevelCols + lNumMetricsCols + 1)
    
    If (Not bExcludeTits) Then lRowOff = 1
    
'   Titles of levels
    If (Not bExcludeTits) Then
        For jj = 0 To lNumLevelCols
            vOut(0, jj) = CStr(vTitles(jj))
        Next jj
'   Titles of metrics
        For jj = lNumLevelCols + 1 To (lNumLevelCols + lNumMetricsCols + 1)
            vOut(0, jj) = CStr(vMetricsTitles(jj - lNumLevelCols - 1))
        Next jj
    End If
'   Levels
    For ii = 0 To lNumSecs - 1 '   Number of secs
        If (sSecurity(ii) <> " Equity") Then  '   I.e. ticker is not empty
            lResRow = dicTickers(Left(sSecurity(ii), Len(sSecurity(ii)) - 7))
            If IsEmpty(vLevels(lResRow, 0)) Then
                vOut(ii + lRowOff, 0) = "No data - check ticker"
            ElseIf (vLevels(lResRow, 21) = "Banks") Or (vLevels(lResRow, 21) = "Insurance") Or (vLevels(lResRow, 17) = "Diversified Finan Serv") Then
                vOut(ii + lRowOff, 0) = vLevels(lResRow, 0) '   Name
                vOut(ii + lRowOff, 1) = "Financials N/A"
            ElseIf vMetrics(lResRow, 15) <= 0 Then
                vOut(ii + lRowOff, 0) = vLevels(lResRow, 0) '   Name
                vOut(ii + lRowOff, 1) = "No sales"
            Else
                If (vLevels(lResRow, 0) <> "") Then
                    For jj = 0 To lNumLevelCols  '   Number of cols
                        vOut(ii + lRowOff, jj) = vLevels(lResRow, jj)
                    Next jj
                End If
            End If
        End If
    Next ii
'   Metrics
    For ii = 0 To lNumSecs - 1 '   Number of secs + title row
        If (sSecurity(ii) <> " Equity") Then  '   I.e. ticker is not empty
            lResRow = dicTickers(Left(sSecurity(ii), Len(sSecurity(ii)) - 7))
            If (vLevels(lResRow, 0) <> "") Then
                For jj = lNumLevelCols + 1 To (lNumLevelCols + lNumMetricsCols + 1)
                    vOut(ii + lRowOff, jj) = vMetrics(lResRow, jj - lNumLevelCols - 1)
                Next jj
            End If
        End If
    Next ii
    
    If bHorizontal Then
        GetFALScores = vOut
    Else
        GetFALScores = WorksheetFunction.Transpose(vOut)
    End If

'   Format cells
'    If (Not bHorizontal) Then   '   Don't format if horizontal
'        With rTickers.Cells(1, 1)
'            For jj = 0 To lNumLevelCols
'                For ii = 0 To lNumSecs - 1
'                    .Offset(ii + 2, jj).Interior.Color = RGB(10, 20, 40) '   HEXCOL2RGB(get_number_color(.Offset(ii + 2, jj).Value))
'
'                Next ii
'            Next jj
'        End With
'    End If
    
    Application.StatusBar = False
Exit Function

shagged:
    MsgBox "Error in GetFALScores"
    'Resume
    Application.StatusBar = False
End Function


Function ReturnFALResults(ByVal vSecurity As Variant, ByVal sYear As String, ByVal sPeriod As String, ByVal sConsolidated As String, ByRef vLevels As Variant, ByRef vMetrics As Variant, ByRef vTitles As Variant, ByRef vMetricsTitles As Variant) As Boolean

Dim vSortedRes As Variant, vSortedRes_Y1 As Variant, vStaticFields As Variant, vRes As Variant, vRes_Y1 As Variant, vInputs As Variant
Dim sOverrideFields(0 To 2) As String, sOverridefields_Y1(0 To 2) As String, sOverrideValues(0 To 2) As String, sOverrideValues_Y1(0 To 2) As String, sSecurity() As String
Dim bOK As Boolean, ii As Long, dicLabels As Dictionary
Dim i As Long, j As Long


    Set dicLabels = New Dictionary
    vTitles = Array("Name", "Level of accounts receivable", "Level of inventory", "Level of other current assets", "Level of accounts payable", "Level of other current liability", "Working capital score", "", "Level of non-operating income as a % recurring income", "Level cashflow from operations as a % net income", "Level freecashflow as a % net income", "Quality of earnings score", "", "Level of debt to ebitda score", "Level of ebitda interest expense score", "Level of debt to OPCF", "Balance sheet score", "", "Overall Score", "")
    vMetricsTitles = Array("Key Financial ratios", "Receivable days", "Inventory days", "Other current asset days", "Payable days", "Other current liability days", "Non-operating as a % recurring income", "Cash from ops as a % net income", "Free cash flow from operations as a % net income", "Debt to Ebitda (x)", "Ebitda interest expense (x)", "Debt to operating cash flow (x)", "Total debt to shareholders equity (%)", "Return on equity (%)", "Key Financial Metrics", "Sales/Revenue/Turnover", "Net Income (Losses)", "Cash from Operations", "Cash from investing activities", "Free Cash Flow", "Total debt")

    bOK = GetFields(vStaticFields)
'    numStaticFields = UBound(vStaticFields) + 1
    
    sOverrideFields(0) = "EQY_FUND_YEAR"
    sOverrideValues(0) = sYear
    sOverrideValues_Y1(0) = CStr(CInt(sYear) - 1)
    sOverrideFields(1) = "FUND_PER"
    sOverrideValues(1) = sPeriod
    sOverrideValues_Y1(1) = sPeriod
    sOverrideFields(2) = "EQY_CONSOLIDATED"
    sOverrideValues(2) = sConsolidated
    sOverrideValues_Y1(2) = sConsolidated

    ReDim sSecurity(LBound(vSecurity) To UBound(vSecurity))
    For ii = LBound(vSecurity) To UBound(vSecurity)
        sSecurity(ii) = vSecurity(ii)
    Next ii

'    sOverrideValues_Y1 = sOverrideValues
'    sOverrideValues_Y1(1) = CStr(CInt(sOverrideValues(1)) - 1)



    bbControlSync.MakeRequest sSecurity, vStaticFields, sOverrideFields, sOverrideValues
    bbControlSync_Y1.MakeRequest sSecurity, vStaticFields, sOverrideFields, sOverrideValues_Y1
    
    Application.StatusBar = "Calculating FAL scores..."
    vRes = bbControlSync.ReturnData '   Get the results back in a matrix
    vRes_Y1 = bbControlSync_Y1.ReturnData '   Get the results back in a matrix
    
    For ii = 0 To UBound(vStaticFields)
        dicLabels.Add Trim(CStr(vStaticFields(ii))), ii '   Put into a dictionary: label name, column
    Next ii
    
    vSortedRes = SortResultsMatrix(dicLabels, vRes) '   Sort into consistent columns
    vSortedRes_Y1 = SortResultsMatrix(dicLabels, vRes_Y1) '   Sort into consistent columns
    
  
'   Get the INPUTS to the scores (e.g. rec. days)
    vInputs = CalcInputs(dicLabels, vSortedRes, vSortedRes_Y1)

'   Calc the SCORES
    vLevels = CalcLevels(vInputs)
    
'   Popupate the financial ratiosand fianncial metrics
    bOK = FinRatios(vInputs, vMetrics)

End Function


Private Function FinRatios(ByVal vInp As Variant, ByRef vOut As Variant) As Boolean

Dim ii As Long

ReDim vOut(0 To UBound(vInp, 1), 0 To 20)

    For ii = 0 To UBound(vInp, 1)
        vOut(ii, 0) = "" '   Blank
        vOut(ii, 1) = vInp(ii, 1) '   Receivable days
        vOut(ii, 2) = vInp(ii, 2) '   Inventory days
        vOut(ii, 3) = vInp(ii, 3) '   Other current asset days
        vOut(ii, 4) = vInp(ii, 5) '   Payable days
        vOut(ii, 5) = vInp(ii, 6) '   Liability days
        vOut(ii, 6) = Max(vInp(ii, 11), -vInp(ii, 11)) '   Non-operating as a % recurring income
        vOut(ii, 7) = vInp(ii, 13) '   Cash from ops as a % net income
        vOut(ii, 8) = vInp(ii, 14) '   FCF as a % net income
        vOut(ii, 9) = vInp(ii, 15) '   Debt 2 EBITDA
        vOut(ii, 10) = vInp(ii, 16) '   Ebitda interest expense
        vOut(ii, 11) = vInp(ii, 17) '   Debt 2 OPCF
        vOut(ii, 12) = vInp(ii, 18) '   Debt 2 Equity
        vOut(ii, 13) = vInp(ii, 19) '   RoE
        vOut(ii, 14) = "" '   Blank
        vOut(ii, 15) = vInp(ii, 20) '   Sales/Revenue/Turnover
        vOut(ii, 16) = vInp(ii, 21) '   Net Income (Losses)
        vOut(ii, 17) = vInp(ii, 22) '   Cash from Operations
        vOut(ii, 18) = vInp(ii, 23) '   Cash from investing activities
        vOut(ii, 19) = vInp(ii, 24) '   Free Cash Flow
        vOut(ii, 20) = vInp(ii, 25) '   Total debt
    Next ii

End Function


Private Function CalcInputs(ByVal dicLabels As Dictionary, ByVal vInputs As Variant, ByVal vInputs_Y1 As Variant) As Variant

Dim lRow As Long, ii As Long
Dim vOut As Variant
Dim dWCDays_Y1() As Double

On Error GoTo shagged

    lRow = UBound(vInputs, 1)
    ReDim vOut(lRow, 0 To 28)
    ReDim dWCDays_Y1(lRow, 6)

    On Error Resume Next    '   If there's a divide by zero error, carry on with the next one
    For ii = 0 To lRow
        vOut(ii, 0) = vInputs(ii, dicLabels("NAME"))
        vOut(ii, 20) = CDbl(vInputs(ii, dicLabels("TRAIL_12M_NET_SALES")))
        vOut(ii, 26) = (vInputs(ii, dicLabels("TICKER_AND_EXCH_CODE")))
        vOut(ii, 27) = (vInputs(ii, dicLabels("INDUSTRY_GROUP")))
        If vInputs(ii, dicLabels("TRAIL_12M_NET_SALES")) > 0 Then
            vOut(ii, 1) = 365 * AvgIfDefined(CDbl(vInputs(ii, dicLabels("BS_ACCT_NOTE_RCV"))), CDbl(vInputs_Y1(ii, dicLabels("BS_ACCT_NOTE_RCV")))) / CDbl(vInputs(ii, dicLabels("TRAIL_12M_NET_SALES")))    '   Receivables days
            vOut(ii, 2) = 365 * AvgIfDefined(CDbl(vInputs(ii, dicLabels("BS_INVENTORIES"))), CDbl(vInputs_Y1(ii, dicLabels("BS_INVENTORIES")))) / CDbl(vInputs(ii, dicLabels("TRAIL_12M_NET_SALES")))    '   Inventory days
            vOut(ii, 3) = 365 * AvgIfDefined(CDbl(vInputs(ii, dicLabels("BS_OTHER_CUR_ASSET"))), CDbl(vInputs_Y1(ii, dicLabels("BS_OTHER_CUR_ASSET")))) / CDbl(vInputs(ii, dicLabels("TRAIL_12M_NET_SALES")))    '   Other current asset days
            vOut(ii, 4) = vOut(ii, 1) + vOut(ii, 2) + vOut(ii, 3)   '   Current asset days
            vOut(ii, 5) = 365 * AvgIfDefined(CDbl(vInputs(ii, dicLabels("BS_ACCT_PAYABLE"))), CDbl(vInputs_Y1(ii, dicLabels("BS_ACCT_PAYABLE")))) / CDbl(vInputs(ii, dicLabels("TRAIL_12M_NET_SALES")))    '   Payable days
            vOut(ii, 6) = 365 * AvgIfDefined(CDbl(vInputs(ii, dicLabels("BS_OTHER_ST_LIAB"))), CDbl(vInputs_Y1(ii, dicLabels("BS_OTHER_ST_LIAB")))) / CDbl(vInputs(ii, dicLabels("TRAIL_12M_NET_SALES")))    '   Other current liability days

            
            vOut(ii, 7) = vOut(ii, 5) + vOut(ii, 6) '   Current liabilities days
            vOut(ii, 8) = (vOut(ii, 1) + vOut(ii, 2) + vOut(ii, 3) + vOut(ii, 5) + vOut(ii, 6))  '   Gross working capital days (A/R, inv, OCA, A/P & OCL)
            vOut(ii, 9) = CDbl(vInputs(ii, dicLabels("IS_OPER_INC"))) - CDbl(vInputs(ii, dicLabels("IS_INT_EXPENSE"))) + CDbl(vInputs(ii, dicLabels("IS_EQY_EARN_FROM_INVEST_ASSOC")))  '   Recurring pre-tax profit (operating profit less interest expense)
            vOut(ii, 10) = CDbl(vInputs(ii, dicLabels("IS_NET_NON_OPER_LOSS"))) + CDbl(vInputs(ii, dicLabels("IS_FOREIGN_EXCH_LOSS"))) + CDbl(vInputs(ii, dicLabels("IS_XO_LOSS_BEF_TAX_EFF"))) - CDbl(vInputs(ii, dicLabels("IS_EQY_EARN_FROM_INVEST_ASSOC"))) '   Non-operating and extraordinary items losses (gains)
'   Quality of Earnings
            vOut(ii, 11) = 100 * vOut(ii, 10) / vOut(ii, 9)   '   Non-operating items as a % normalised pre-tax profit
            vOut(ii, 12) = PosVal(CDbl(vOut(ii, 11))) '   Non-operating as a % normalised pre-tax (adjusted)
            'vOut(ii, 13) = 100 * CDbl(vInputs(ii, dicLabels("CF_CASH_FROM_OPER"))) / CDbl(vInputs(ii, dicLabels("NET_INCOME"))) '   Cashflow from operations as a % net income
            If IsEmpty(vInputs(ii, dicLabels("CF_CASH_FROM_OPER"))) Then
                vOut(ii, 13) = "na"
'            ElseIf (vInputs(ii, dicLabels("CF_CASH_FROM_OPER"))) < 0 Then
'                vOut(ii, 13) = "na"
            Else
                vOut(ii, 13) = 100 * CDbl(vInputs(ii, dicLabels("CF_CASH_FROM_OPER"))) / CDbl(vInputs(ii, dicLabels("NET_INCOME"))) '   Cashflow from operations as a % net income
            End If
            'vOut(ii, 14) = 100 * (CDbl(vInputs(ii, dicLabels("CF_CASH_FROM_INV_ACT"))) + CDbl(vInputs(ii, dicLabels("CF_CASH_FROM_OPER")))) / CDbl(vInputs(ii, dicLabels("NET_INCOME")))     '   FCF as a % net income
            If IsEmpty(vInputs(ii, dicLabels("CF_CASH_FROM_INV_ACT"))) Then
                vOut(ii, 14) = "na"
            Else
                vOut(ii, 14) = 100 * (CDbl(vInputs(ii, dicLabels("CF_CASH_FROM_INV_ACT"))) + CDbl(vInputs(ii, dicLabels("CF_CASH_FROM_OPER")))) / CDbl(vInputs(ii, dicLabels("NET_INCOME")))     '   FCF as a % net income
            End If
'   Balance sheet
            If IsEmpty(vInputs(ii, dicLabels("TOT_DEBT_TO_EBITDA"))) Then
                vOut(ii, 15) = "na"
            Else
                vOut(ii, 15) = CDbl(vInputs(ii, dicLabels("TOT_DEBT_TO_EBITDA")))     '     Debt 2 EBITDA
            End If
            vOut(ii, 16) = CDbl(vInputs(ii, dicLabels("EBITDA_TO_TOT_INT_EXP")))    '   Ebitda interest expense
            'vOut(ii, 17) = CDbl(vInputs(ii, dicLabels("SHORT_AND_LONG_TERM_DEBT"))) / CDbl(vInputs(ii, dicLabels("CF_CASH_FROM_OPER")))  '   Level of Debt / OPCF
            If IsEmpty(vInputs(ii, dicLabels("CF_CASH_FROM_OPER"))) Then
                vOut(ii, 17) = "na"
            Else
                vOut(ii, 17) = CDbl(vInputs(ii, dicLabels("SHORT_AND_LONG_TERM_DEBT"))) / CDbl(vInputs(ii, dicLabels("CF_CASH_FROM_OPER")))  '   Level of Debt / OPCF
            End If
            vOut(ii, 18) = 100 * CDbl(vInputs(ii, dicLabels("SHORT_AND_LONG_TERM_DEBT"))) / CDbl(vInputs(ii, dicLabels("TOT_SHRHLDR_EQY"))) '   D2E
            vOut(ii, 19) = 200 * CDbl(vInputs(ii, dicLabels("NET_INCOME"))) / (CDbl(vInputs(ii, dicLabels("TOT_SHRHLDR_EQY"))) + CDbl(vInputs_Y1(ii, dicLabels("TOT_SHRHLDR_EQY")))) '
            vOut(ii, 21) = CDbl(vInputs(ii, dicLabels("NET_INCOME")))
            vOut(ii, 22) = CDbl(vInputs(ii, dicLabels("CF_CASH_FROM_OPER")))
            vOut(ii, 23) = CDbl(vInputs(ii, dicLabels("CF_CASH_FROM_INV_ACT")))
            vOut(ii, 24) = (CDbl(vInputs(ii, dicLabels("CF_CASH_FROM_INV_ACT"))) + CDbl(vInputs(ii, dicLabels("CF_CASH_FROM_OPER"))))   '   FCF
            vOut(ii, 25) = CDbl(vInputs(ii, dicLabels("SHORT_AND_LONG_TERM_DEBT")))
            vOut(ii, 28) = CDbl(vInputs(ii, dicLabels("IS_INT_EXPENSE")))
'   Others
        End If
    Next ii

    CalcInputs = vOut
    
Exit Function

shagged:

    MsgBox "Error in CalcInputs"
    'Resume
End Function


Private Function CalcLevels(ByVal vInputs As Variant) As Variant

Dim lRow As Long, ii As Long, jj As Long
Dim vOut As Variant, vLookupArray As Variant, vSelAvg As Variant, vGrossCap As Variant
Dim dRecDaysScore As Double, dRawWCScore() As Double, dRawQoEScore() As Double, dRawBSScore() As Double, dRawOverallScore() As Double

On Error GoTo shagged

    lRow = UBound(vInputs, 1)
    ReDim vOut(lRow, 22)
    ReDim dRawWCScore(lRow)
    ReDim dRawQoEScore(lRow)
    ReDim dRawBSScore(lRow)
    ReDim dRawOverallScore(lRow)
    
    For ii = 0 To lRow
        For jj = 0 To UBound(vOut, 2)
            vOut(ii, jj) = "na"
        Next jj
    Next ii

    For ii = 0 To lRow
        vOut(ii, 0) = vInputs(ii, 0) '   Name
        vOut(ii, 1) = 10 * SimpleLookUp(vInputs(ii, 1), LevelOfAccountsReceivable, 2) '   Receivable level
        vOut(ii, 2) = 10 * SimpleLookUp(vInputs(ii, 2), LevelOfInventory, 2) '   Inventory level
        vOut(ii, 3) = 10 * SimpleLookUp(vInputs(ii, 3), LevelOfOtherCurrentAssets, 2) '   Other current assets level
        vOut(ii, 4) = 10 * SimpleLookUp(vInputs(ii, 5), LevelOfAccoutsPayable, 2) '   Payable level
        vOut(ii, 5) = 10 * SimpleLookUp(vInputs(ii, 6), LevelOfOtherCurrentLiabilities, 2) '   Other current liability level
        vGrossCap = 10 * SimpleLookUp(vInputs(ii, 8), LevelOfGrossWorkingCapital, 2)  ' Gross working capital level
        vSelAvg = SelectiveAverage(Array(vOut(ii, 1), vOut(ii, 2), vOut(ii, 3), vOut(ii, 4), vOut(ii, 5), vGrossCap))
        If IsNumeric(vSelAvg) Then
            dRawWCScore(ii) = vSelAvg
            vOut(ii, 6) = AdjustedScore(dRawWCScore(ii), WorkingCapital)  '   WC score
        End If
        vOut(ii, 7) = ""
'   Q of E scores
        vOut(ii, 8) = 10 * SimpleLookUp(vInputs(ii, 12), LevelOfNonOperatingIncomeAsAPCRecurringIncome, 2) ' Level Non Operating Income / recurring income
        If (vInputs(ii, 13) = "na") Then
            vOut(ii, 9) = "na"
        ElseIf (vInputs(ii, 21) < 0) Then
            vOut(ii, 9) = 100   '   If the NET INCOME is -ve, cf from ops always goes to 100.  TS 7 Nov 2013
        Else
            vOut(ii, 9) = 100 - 10 * SimpleLookUp(vInputs(ii, 13), LevelOfCashflowFromOperationsAsAPCNetIncome, 2) ' Level Cash Flow from Operations / Net income
        End If
        If (vInputs(ii, 14) = "na") Then
            vOut(ii, 10) = 100
        ElseIf (vInputs(ii, 21) < 0) Then
            vOut(ii, 10) = 100
        Else
            vOut(ii, 10) = 100 - 10 * SimpleLookUp(vInputs(ii, 14), LevelOfFreeCashFlowAsAPCNetIncome, 2) ' Level Free Cash Flow / Net income
        End If
        vSelAvg = SelectiveAverage(Array(vOut(ii, 8), vOut(ii, 9), vOut(ii, 10)))
        If IsNumeric(vSelAvg) Then
            dRawQoEScore(ii) = vSelAvg
            vOut(ii, 11) = AdjustedScore(dRawQoEScore(ii), QualityOfEarnings) ' QofE score
        End If
        vOut(ii, 12) = ""
'   B S scores
        If vInputs(ii, 15) = "na" Then '   "TOT_DEBT_TO_EBITDA"
            vOut(ii, 13) = 100 ' "na"   '   TS  19 Nov 2013
        ElseIf vInputs(ii, 15) <= 0 Then
            vOut(ii, 13) = 0
        Else
            vOut(ii, 13) = 10 * SimpleLookUp(vInputs(ii, 15), LevelOfDebtToEbitdaScore, 2) '   Level D 2 EBITDA
        End If
        If vInputs(ii, 28) <= 0 Then    '   "IS_INT_EXPENSE"
            vOut(ii, 14) = 0
        Else
            vOut(ii, 14) = 100 - 10 * SimpleLookUp(vInputs(ii, 16), LevelOfEbitdaInterestExpense, 2) ' Level of ebitda interest expense score
        End If
        If (vInputs(ii, 17) = "na") Then    '   "SHORT_AND_LONG_TERM_DEBT" / "CF_CASH_FROM_OPER"
            vOut(ii, 15) = "na"
        ElseIf vInputs(ii, 17) < 0 Then
            vOut(ii, 15) = 100
        Else
            vOut(ii, 15) = 10 * SimpleLookUp(vInputs(ii, 17), LevelOfDebt2OPCF, 2) '   Level D2OPCF
        End If
        vSelAvg = SelectiveAverage(Array(vOut(ii, 13), vOut(ii, 14), vOut(ii, 15)))
        If IsNumeric(vSelAvg) Then
            dRawBSScore(ii) = vSelAvg
            vOut(ii, 16) = AdjustedScore(dRawBSScore(ii), BalanceSheet) '   BS Score
        End If
        vOut(ii, 17) = ""
'   Overall score
        vOut(ii, 18) = AdjustedScore(SelectiveAverage(Array(dRawWCScore(ii), dRawQoEScore(ii), dRawBSScore(ii))), Overall)
        vOut(ii, 19) = ""
        vOut(ii, 20) = vInputs(ii, 26) '   Ticker
        vOut(ii, 21) = vInputs(ii, 27) '   INDUSTRY_GROUP
        vOut(ii, 22) = ""
    Next ii

    CalcLevels = vOut
    
Exit Function

shagged:

    MsgBox "Error in CalcLevels"
    'Resume
End Function


Private Function SelectiveAverage(ByVal vIn As Variant) As Variant

Dim ii As Long, lLen As Long, lCount As Long
Dim dOut As Double, dCum As Double
On Error GoTo shagged
    lLen = UBound(vIn)
    For ii = LBound(vIn) To UBound(vIn)
        If vIn(ii) = "na" Then
        Else
            dCum = dCum + CDbl(vIn(ii))
            lCount = lCount + 1
        End If
    Next ii
    
    If lCount > 0 Then
        SelectiveAverage = dCum / lCount
    Else
        SelectiveAverage = "na"
    End If
    
Exit Function

shagged:
    MsgBox "Error in SelectiveAverage"
    'Resume
End Function



Private Function SortResultsMatrix(ByVal dicLabels As Dictionary, ByVal vInputs As Variant) As Variant

Dim sFirst As String, sLast As String, sSplit() As String, sFields() As String
Dim lRow As Long, lCol As Long, ii As Long, lInputCol As Long

Dim vOut As Variant

On Error GoTo shagged

    ReDim vOut(UBound(vInputs, 1), UBound(vInputs, 2))
'    For lRow = 0 To UBound(vInputs, 1) - 1
'        For lCol = 0 To UBound(vInputs, 2) - 1
'            vOut(lRow, lCol) = "0.0"  '   Initialise with 0 as otherwise there's problems with the arithmatic later on.
'        Next lCol
'    Next lRow
    ReDim sFields(dicLabels.Count - 1)
    For lRow = 0 To UBound(vInputs, 1)
        For lCol = 0 To UBound(vInputs, 2) - 1
            If vInputs(lRow, lCol) = "" Then
                ii = ii + 1
            Else
                sSplit = Split(vInputs(lRow, lCol), " = ")
                sFirst = sSplit(0)
                sLast = CStr(sSplit(1))
            End If
            If dicLabels.Exists(sFirst) Then
                lInputCol = dicLabels(sFirst)
                vOut(lRow, lInputCol) = Trim(sLast)
            End If
        Next lCol
    Next lRow

    SortResultsMatrix = vOut

Exit Function

shagged:

    MsgBox "Error in SortResultsMatrix"
    'Resume
End Function


Function SimpleLookUp(ByVal dLookupValue As Double, ByVal vArray As Variant, ByVal iCol As Integer) As Double
    '   vArray must be ordered N x N array of doubles
Dim lRows As Long, ii As Long
Dim dOut As Double, dLower As Double, dUpper As Double, dFrac As Double

On Error GoTo shagged

    lRows = UBound(vArray, 1)

    If dLookupValue < vArray(1, 1) Then
        dOut = vArray(1, 2)
    ElseIf dLookupValue > vArray(lRows, 1) Then
        dOut = vArray(lRows, 2)
    Else
        If dLookupValue = 0 Then
            dOut = 0
        Else
            For ii = 1 To lRows
                If vArray(ii, 1) >= dLookupValue Then
                    dLower = vArray(ii - 1, iCol)
                    dUpper = vArray(ii, iCol)
                    dFrac = (dLookupValue - vArray(ii - 1, 1)) / (vArray(ii, 1) - vArray(ii - 1, 1))
                    dOut = (dFrac * dUpper) + ((1 - dFrac) * dLower)
                    Exit For
                End If
            Next ii
        End If
    End If
    
    SimpleLookUp = dOut

Exit Function

shagged:

    MsgBox "Error in SimpleLookUp"
    'Resume
End Function



Sub RWLookups()

Dim r As Range
Dim ii As Integer

    Set r = Range("RStart")
    For ii = 1 To 100
        Debug.Print "d(" & ii & ", 1) = " & r.Offset(ii - 1, 0).Value
    Next ii
    For ii = 1 To 100
        Debug.Print "d(" & ii & ", 2) = " & r.Offset(ii - 1, 1).Value
    Next ii

End Sub

Function PosVal(d As Double) As Double
    PosVal = d
    If d < 0 Then
        PosVal = -d
    End If
End Function


Function GetFields(ByRef v As Variant) As Boolean

v = Array("TICKER_AND_EXCH_CODE", "NAME", "PX_LAST", "MARKET_SECTOR_DES", "HISTORICAL_MARKET_CAP", "TRAIL_12M_NET_SALES", "IS_COGS_TO_FE_AND_PP_AND_G", _
"BS_ACCT_NOTE_RCV", "BS_OTHER_CUR_ASSET", "BS_INVENTORIES", "BS_ACCT_PAYABLE", "BS_OTHER_ST_LIAB", "ACCT_RCV_DAYS", _
"IS_OPER_INC", "IS_INT_EXPENSE", "EFF_INT_RATE", "IS_CAP_INT_EXP", "IS_EQY_EARN_FROM_INVEST_ASSOC", "IS_NET_NON_OPER_LOSS", _
"IS_FOREIGN_EXCH_LOSS", "IS_XO_LOSS_BEF_TAX_EFF", "CF_CASH_FROM_OPER", "NET_INCOME", "CF_CASH_FROM_INV_ACT", _
"EBITDA_TO_TOT_INT_EXP", "TOT_DEBT_TO_EBITDA", "SHORT_AND_LONG_TERM_DEBT", "EBITDA", "TOT_SHRHLDR_EQY", _
"CAPITAL_EMPLOYED", "CF_DVD_PAID", "CF_INCR_CAP_STOCK", "INDUSTRY_GROUP")

GetFields = True

End Function

Function Max(d1, d2) As Variant
    Max = d1
    If CDbl(d2) > CDbl(d1) Then Max = d2
End Function

Function AvgIfDefined(s1 As Double, s2 As Double) As Double
'   Problem is if a value is undefined for a particular year, don't average it
Dim r As Double
    If ((s1 = 0) And (s2 = 0)) Then
        r = 0
    ElseIf (s1 = 0) Then
        r = s2
    ElseIf (s2 = 0) Then
        r = s1
    Else
        r = (s1 + s2) / 2
    End If
    AvgIfDefined = r
End Function

'Function AvgIfBothDefined(s1 As Variant, s2 As Variant) As Double
''   Returns value if both are defined, else NA
'Dim r As Double
'    If ((s1 = 0) And (s2 = 0)) Then
'        r = 0
'    ElseIf (s1 = 0) Then
'        r = s2
'    ElseIf (s2 = 0) Then
'        r = s1
'    Else
'        r = (s1 + s2) / 2
'    End If
'    AvgIfDefined = r
'End Function

Function Version() As String
    Version = "1.3.1"
End Function

