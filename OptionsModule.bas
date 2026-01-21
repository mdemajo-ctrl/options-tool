Attribute VB_Name = "OptionsModule"
Option Explicit

' Constants for sheet layout - Input area
Private Const INPUT_ROW As Long = 2
Private Const TICKER_COL As Long = 2        ' Column B
Private Const EXPIRY_COL As Long = 3        ' Column C
Private Const OPTTYPE_COL As Long = 4       ' Column D
Private Const STOCK_PRICE_COL As Long = 6   ' Column F
Private Const STOCK_AT_EXPIRY_COL As Long = 8  ' Column H

' Constants for P&L summary area (Row 3)
Private Const SUMMARY_ROW As Long = 3
Private Const TOTAL_PNL_COL As Long = 12    ' Column L

' Constants for options data grid
Private Const DATA_START_ROW As Long = 6
Private Const STRIKE_COL As Long = 1        ' Column A
Private Const BID_COL As Long = 2           ' Column B
Private Const ASK_COL As Long = 3           ' Column C
Private Const LAST_COL As Long = 4          ' Column D
Private Const MID_COL As Long = 5           ' Column E
Private Const VOLUME_COL As Long = 6        ' Column F
Private Const OI_COL As Long = 7            ' Column G
Private Const IV_COL As Long = 8            ' Column H
Private Const DELTA_COL As Long = 9         ' Column I
Private Const POSITION_COL As Long = 10     ' Column J
Private Const ENTRY_PRICE_COL As Long = 11  ' Column K
Private Const VALUE_AT_EXPIRY_COL As Long = 12  ' Column L
Private Const PNL_COL As Long = 13          ' Column M

Public Sub FetchOptionsData()
    Dim ws As Worksheet
    Dim ticker As String
    Dim expiryDate As Date
    Dim optionType As String
    Dim url As String
    Dim response As String
    Dim unixTimestamp As Double

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    ' Read input parameters
    ticker = Trim(UCase(ws.Cells(INPUT_ROW, TICKER_COL).Value))
    optionType = Trim(UCase(ws.Cells(INPUT_ROW, OPTTYPE_COL).Value))

    ' Validate inputs
    If ticker = "" Then
        MsgBox "Please enter a stock ticker symbol.", vbExclamation
        Exit Sub
    End If

    If optionType <> "CALL" And optionType <> "PUT" Then
        MsgBox "Please enter CALL or PUT for option type.", vbExclamation
        Exit Sub
    End If

    ' Handle expiry date
    If IsDate(ws.Cells(INPUT_ROW, EXPIRY_COL).Value) Then
        expiryDate = ws.Cells(INPUT_ROW, EXPIRY_COL).Value
        unixTimestamp = DateToUnix(expiryDate)
    Else
        ' If no date provided, we'll get the nearest expiration
        unixTimestamp = 0
    End If

    ' Clear previous data (but preserve positions if strikes match)
    ClearOptionsData ws

    ' Update status
    Application.StatusBar = "Fetching options data for " & ticker & "..."

    ' Build URL
    If unixTimestamp > 0 Then
        url = "https://query1.finance.yahoo.com/v7/finance/options/" & ticker & "?date=" & CStr(CLng(unixTimestamp))
    Else
        url = "https://query1.finance.yahoo.com/v7/finance/options/" & ticker
    End If

    ' Fetch data
    response = HttpGet(url)

    If response = "" Then
        MsgBox "Failed to fetch data. Please check your internet connection and ticker symbol.", vbExclamation
        GoTo Cleanup
    End If

    ' Parse and display data
    ParseAndDisplayOptions ws, response, optionType

    ' Add P&L formulas
    AddPnLFormulas ws

    Application.StatusBar = "Options data loaded successfully."

Cleanup:
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub FetchExpirationDates()
    Dim ws As Worksheet
    Dim ticker As String
    Dim url As String
    Dim response As String
    Dim expirations As String
    Dim msg As String

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet
    ticker = Trim(UCase(ws.Cells(INPUT_ROW, TICKER_COL).Value))

    If ticker = "" Then
        MsgBox "Please enter a stock ticker symbol first.", vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "Fetching expiration dates for " & ticker & "..."

    url = "https://query1.finance.yahoo.com/v7/finance/options/" & ticker
    response = HttpGet(url)

    If response = "" Then
        MsgBox "Failed to fetch data.", vbExclamation
        GoTo Cleanup
    End If

    ' Extract expiration dates
    expirations = ExtractExpirationDates(response)

    If expirations <> "" Then
        msg = "Available expiration dates for " & ticker & ":" & vbCrLf & vbCrLf & expirations
        MsgBox msg, vbInformation, "Expiration Dates"
    Else
        MsgBox "No expiration dates found.", vbExclamation
    End If

Cleanup:
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub CalculatePnL()
    ' Recalculate P&L - useful after changing stock at expiry or positions
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Force recalculation
    ws.Calculate

    MsgBox "P&L recalculated.", vbInformation
End Sub

Private Function HttpGet(url As String) As String
    Dim http As Object

    On Error GoTo ErrorHandler

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    http.send

    If http.Status = 200 Then
        HttpGet = http.responseText
    Else
        HttpGet = ""
    End If

    Exit Function

ErrorHandler:
    HttpGet = ""
End Function

Private Sub ClearOptionsData(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, STRIKE_COL).End(xlUp).Row
    If lastRow >= DATA_START_ROW Then
        ' Clear market data columns but preserve position and entry price
        ws.Range(ws.Cells(DATA_START_ROW, STRIKE_COL), ws.Cells(lastRow, IV_COL)).ClearContents
        ws.Range(ws.Cells(DATA_START_ROW, DELTA_COL), ws.Cells(lastRow, DELTA_COL)).ClearContents
        ws.Range(ws.Cells(DATA_START_ROW, VALUE_AT_EXPIRY_COL), ws.Cells(lastRow, PNL_COL)).ClearContents
    End If
    ws.Cells(INPUT_ROW, STOCK_PRICE_COL).ClearContents
End Sub

Private Sub ParseAndDisplayOptions(ws As Worksheet, jsonResponse As String, optionType As String)
    Dim json As Object
    Dim optionChain As Object
    Dim result As Object
    Dim quote As Object
    Dim options As Object
    Dim contracts As Object
    Dim contract As Object
    Dim row As Long
    Dim i As Long
    Dim bid As Double, ask As Double

    On Error GoTo ErrorHandler

    ' Parse JSON
    Set json = JsonConverter.ParseJson(jsonResponse)

    ' Navigate to option chain
    Set optionChain = json("optionChain")
    Set result = optionChain("result")(1)

    ' Get stock price
    Set quote = result("quote")
    ws.Cells(INPUT_ROW, STOCK_PRICE_COL).Value = quote("regularMarketPrice")

    ' Set default stock at expiry to current price if empty
    If ws.Cells(INPUT_ROW, STOCK_AT_EXPIRY_COL).Value = "" Then
        ws.Cells(INPUT_ROW, STOCK_AT_EXPIRY_COL).Value = quote("regularMarketPrice")
    End If

    ' Get options
    Set options = result("options")(1)

    ' Select calls or puts
    If optionType = "CALL" Then
        Set contracts = options("calls")
    Else
        Set contracts = options("puts")
    End If

    ' Write headers
    WriteHeaders ws

    ' Write option data
    row = DATA_START_ROW
    For i = 1 To contracts.Count
        Set contract = contracts(i)

        ws.Cells(row, STRIKE_COL).Value = GetJsonValue(contract, "strike")

        bid = GetJsonValue(contract, "bid")
        ask = GetJsonValue(contract, "ask")

        ws.Cells(row, BID_COL).Value = bid
        ws.Cells(row, ASK_COL).Value = ask
        ws.Cells(row, LAST_COL).Value = GetJsonValue(contract, "lastPrice")

        ' Calculate mid price
        If bid > 0 And ask > 0 Then
            ws.Cells(row, MID_COL).Value = (bid + ask) / 2
        Else
            ws.Cells(row, MID_COL).Value = GetJsonValue(contract, "lastPrice")
        End If

        ws.Cells(row, VOLUME_COL).Value = GetJsonValue(contract, "volume")
        ws.Cells(row, OI_COL).Value = GetJsonValue(contract, "openInterest")
        ws.Cells(row, IV_COL).Value = GetJsonValue(contract, "impliedVolatility")

        ' Get delta if available (Yahoo doesn't always provide this)
        Dim deltaVal As Variant
        deltaVal = GetJsonValueOrNull(contract, "delta")
        If Not IsNull(deltaVal) Then
            ws.Cells(row, DELTA_COL).Value = deltaVal
        End If

        ' Initialize position to 0 if empty
        If ws.Cells(row, POSITION_COL).Value = "" Then
            ws.Cells(row, POSITION_COL).Value = 0
        End If

        ' Set entry price to mid if empty and position is non-zero
        If ws.Cells(row, ENTRY_PRICE_COL).Value = "" Then
            ws.Cells(row, ENTRY_PRICE_COL).Value = ws.Cells(row, MID_COL).Value
        End If

        row = row + 1
    Next i

    ' Format the data
    FormatOptionsData ws, row - 1

    Exit Sub

ErrorHandler:
    MsgBox "Error parsing options data: " & Err.Description, vbCritical
End Sub

Private Sub AddPnLFormulas(ws As Worksheet)
    Dim lastRow As Long
    Dim row As Long
    Dim optionType As String
    Dim stockAtExpiryCell As String

    lastRow = ws.Cells(ws.Rows.Count, STRIKE_COL).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub

    optionType = Trim(UCase(ws.Cells(INPUT_ROW, OPTTYPE_COL).Value))
    stockAtExpiryCell = "$" & ColLetter(STOCK_AT_EXPIRY_COL) & "$" & INPUT_ROW

    For row = DATA_START_ROW To lastRow
        Dim strikeCell As String
        Dim positionCell As String
        Dim entryPriceCell As String
        Dim valueAtExpiryCell As String

        strikeCell = ColLetter(STRIKE_COL) & row
        positionCell = ColLetter(POSITION_COL) & row
        entryPriceCell = ColLetter(ENTRY_PRICE_COL) & row
        valueAtExpiryCell = ColLetter(VALUE_AT_EXPIRY_COL) & row

        ' Value at expiry formula
        ' For CALL: MAX(StockAtExpiry - Strike, 0)
        ' For PUT: MAX(Strike - StockAtExpiry, 0)
        If optionType = "CALL" Then
            ws.Cells(row, VALUE_AT_EXPIRY_COL).Formula = "=MAX(" & stockAtExpiryCell & "-" & strikeCell & ",0)"
        Else
            ws.Cells(row, VALUE_AT_EXPIRY_COL).Formula = "=MAX(" & strikeCell & "-" & stockAtExpiryCell & ",0)"
        End If

        ' P&L formula: (ValueAtExpiry - EntryPrice) * Position * 100
        ws.Cells(row, PNL_COL).Formula = "=(" & valueAtExpiryCell & "-" & entryPriceCell & ")*" & positionCell & "*100"
    Next row

    ' Add total P&L formula
    ws.Cells(SUMMARY_ROW, TOTAL_PNL_COL).Formula = "=SUM(" & ColLetter(PNL_COL) & DATA_START_ROW & ":" & ColLetter(PNL_COL) & lastRow & ")"
    ws.Cells(SUMMARY_ROW, TOTAL_PNL_COL - 1).Value = "Total P&L:"
    ws.Cells(SUMMARY_ROW, TOTAL_PNL_COL - 1).Font.Bold = True

    ' Format total P&L
    With ws.Cells(SUMMARY_ROW, TOTAL_PNL_COL)
        .NumberFormat = "$#,##0.00"
        .Font.Bold = True
    End With

    ' Add conditional formatting to Total P&L (green if positive, red if negative)
    With ws.Cells(SUMMARY_ROW, TOTAL_PNL_COL)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(1).Font.Color = RGB(0, 128, 0)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(2).Font.Color = RGB(192, 0, 0)
    End With
End Sub

Private Function ColLetter(colNum As Long) As String
    Dim n As Long
    Dim c As String
    Dim s As String

    n = colNum
    Do
        c = Chr(((n - 1) Mod 26) + Asc("A"))
        s = c & s
        n = (n - 1) \ 26
    Loop While n > 0

    ColLetter = s
End Function

Private Sub WriteHeaders(ws As Worksheet)
    ws.Cells(DATA_START_ROW - 1, STRIKE_COL).Value = "Strike"
    ws.Cells(DATA_START_ROW - 1, BID_COL).Value = "Bid"
    ws.Cells(DATA_START_ROW - 1, ASK_COL).Value = "Ask"
    ws.Cells(DATA_START_ROW - 1, LAST_COL).Value = "Last"
    ws.Cells(DATA_START_ROW - 1, MID_COL).Value = "Mid"
    ws.Cells(DATA_START_ROW - 1, VOLUME_COL).Value = "Volume"
    ws.Cells(DATA_START_ROW - 1, OI_COL).Value = "Open Int"
    ws.Cells(DATA_START_ROW - 1, IV_COL).Value = "Impl Vol"
    ws.Cells(DATA_START_ROW - 1, DELTA_COL).Value = "Delta"
    ws.Cells(DATA_START_ROW - 1, POSITION_COL).Value = "Position"
    ws.Cells(DATA_START_ROW - 1, ENTRY_PRICE_COL).Value = "Entry Price"
    ws.Cells(DATA_START_ROW - 1, VALUE_AT_EXPIRY_COL).Value = "Value@Exp"
    ws.Cells(DATA_START_ROW - 1, PNL_COL).Value = "P&L"

    ' Format headers - market data
    With ws.Range(ws.Cells(DATA_START_ROW - 1, STRIKE_COL), ws.Cells(DATA_START_ROW - 1, DELTA_COL))
        .Font.Bold = True
        .Interior.Color = RGB(68, 84, 106)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Format headers - position/trading columns (different color)
    With ws.Range(ws.Cells(DATA_START_ROW - 1, POSITION_COL), ws.Cells(DATA_START_ROW - 1, PNL_COL))
        .Font.Bold = True
        .Interior.Color = RGB(79, 129, 189)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

Private Sub FormatOptionsData(ws As Worksheet, lastDataRow As Long)
    If lastDataRow < DATA_START_ROW Then Exit Sub

    ' Strike column
    With ws.Range(ws.Cells(DATA_START_ROW, STRIKE_COL), ws.Cells(lastDataRow, STRIKE_COL))
        .NumberFormat = "$#,##0.00"
    End With

    ' Price columns (Bid, Ask, Last, Mid)
    With ws.Range(ws.Cells(DATA_START_ROW, BID_COL), ws.Cells(lastDataRow, MID_COL))
        .NumberFormat = "$#,##0.00"
    End With

    ' Volume and OI
    With ws.Range(ws.Cells(DATA_START_ROW, VOLUME_COL), ws.Cells(lastDataRow, OI_COL))
        .NumberFormat = "#,##0"
    End With

    ' Implied volatility
    With ws.Range(ws.Cells(DATA_START_ROW, IV_COL), ws.Cells(lastDataRow, IV_COL))
        .NumberFormat = "0.00%"
    End With

    ' Delta
    With ws.Range(ws.Cells(DATA_START_ROW, DELTA_COL), ws.Cells(lastDataRow, DELTA_COL))
        .NumberFormat = "0.00"
    End With

    ' Position (integer)
    With ws.Range(ws.Cells(DATA_START_ROW, POSITION_COL), ws.Cells(lastDataRow, POSITION_COL))
        .NumberFormat = "0"
        .Interior.Color = RGB(255, 255, 200)  ' Light yellow to indicate editable
    End With

    ' Entry price
    With ws.Range(ws.Cells(DATA_START_ROW, ENTRY_PRICE_COL), ws.Cells(lastDataRow, ENTRY_PRICE_COL))
        .NumberFormat = "$#,##0.00"
        .Interior.Color = RGB(255, 255, 200)  ' Light yellow to indicate editable
    End With

    ' Value at expiry
    With ws.Range(ws.Cells(DATA_START_ROW, VALUE_AT_EXPIRY_COL), ws.Cells(lastDataRow, VALUE_AT_EXPIRY_COL))
        .NumberFormat = "$#,##0.00"
    End With

    ' P&L with conditional formatting
    With ws.Range(ws.Cells(DATA_START_ROW, PNL_COL), ws.Cells(lastDataRow, PNL_COL))
        .NumberFormat = "$#,##0.00"
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(1).Font.Color = RGB(0, 128, 0)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(2).Font.Color = RGB(192, 0, 0)
    End With

    ' Auto-fit columns
    ws.Columns("A:M").AutoFit
End Sub

Private Function GetJsonValue(obj As Object, key As String) As Variant
    On Error Resume Next
    GetJsonValue = obj(key)
    If Err.Number <> 0 Then GetJsonValue = 0
    On Error GoTo 0
End Function

Private Function GetJsonValueOrNull(obj As Object, key As String) As Variant
    On Error Resume Next
    GetJsonValueOrNull = obj(key)
    If Err.Number <> 0 Then GetJsonValueOrNull = Null
    On Error GoTo 0
End Function

Private Function DateToUnix(dt As Date) As Double
    DateToUnix = (dt - DateSerial(1970, 1, 1)) * 86400
End Function

Private Function UnixToDate(unixTime As Double) As Date
    UnixToDate = DateSerial(1970, 1, 1) + (unixTime / 86400)
End Function

Private Function ExtractExpirationDates(jsonResponse As String) As String
    Dim json As Object
    Dim optionChain As Object
    Dim result As Object
    Dim expirations As Object
    Dim i As Long
    Dim dates As String
    Dim expDate As Date

    On Error GoTo ErrorHandler

    Set json = JsonConverter.ParseJson(jsonResponse)
    Set optionChain = json("optionChain")
    Set result = optionChain("result")(1)
    Set expirations = result("expirationDates")

    dates = ""
    For i = 1 To expirations.Count
        expDate = UnixToDate(expirations(i))
        dates = dates & Format(expDate, "yyyy-mm-dd") & vbCrLf
    Next i

    ExtractExpirationDates = dates
    Exit Function

ErrorHandler:
    ExtractExpirationDates = ""
End Function

Public Sub SetupSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Clear the sheet
    ws.Cells.Clear

    ' Set up input labels - Row 1
    ws.Cells(1, TICKER_COL).Value = "Ticker"
    ws.Cells(1, EXPIRY_COL).Value = "Expiry (YYYY-MM-DD)"
    ws.Cells(1, OPTTYPE_COL).Value = "Type (CALL/PUT)"
    ws.Cells(1, STOCK_PRICE_COL).Value = "Current Price"
    ws.Cells(1, STOCK_AT_EXPIRY_COL).Value = "Stock @ Expiry"

    ' Format input labels
    With ws.Range(ws.Cells(1, TICKER_COL), ws.Cells(1, OPTTYPE_COL))
        .Font.Bold = True
        .Interior.Color = RGB(68, 84, 106)
        .Font.Color = RGB(255, 255, 255)
    End With

    With ws.Cells(1, STOCK_PRICE_COL)
        .Font.Bold = True
        .Interior.Color = RGB(68, 84, 106)
        .Font.Color = RGB(255, 255, 255)
    End With

    With ws.Cells(1, STOCK_AT_EXPIRY_COL)
        .Font.Bold = True
        .Interior.Color = RGB(79, 129, 189)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Set default values
    ws.Cells(INPUT_ROW, TICKER_COL).Value = "AAPL"
    ws.Cells(INPUT_ROW, OPTTYPE_COL).Value = "CALL"

    ' Format stock at expiry cell as editable
    With ws.Cells(INPUT_ROW, STOCK_AT_EXPIRY_COL)
        .Interior.Color = RGB(255, 255, 200)
        .NumberFormat = "$#,##0.00"
    End With

    ' Add data validation for option type
    With ws.Cells(INPUT_ROW, OPTTYPE_COL).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="CALL,PUT"
    End With

    ' Add borders to input area
    With ws.Range(ws.Cells(1, TICKER_COL), ws.Cells(INPUT_ROW, OPTTYPE_COL))
        .Borders.LineStyle = xlContinuous
    End With

    ws.Cells(1, STOCK_PRICE_COL).Borders.LineStyle = xlContinuous
    ws.Cells(INPUT_ROW, STOCK_PRICE_COL).Borders.LineStyle = xlContinuous
    ws.Cells(1, STOCK_AT_EXPIRY_COL).Borders.LineStyle = xlContinuous
    ws.Cells(INPUT_ROW, STOCK_AT_EXPIRY_COL).Borders.LineStyle = xlContinuous

    ' Auto-fit columns
    ws.Columns("A:M").AutoFit

    MsgBox "Sheet setup complete!" & vbCrLf & vbCrLf & _
           "INPUT AREA:" & vbCrLf & _
           "1. Enter ticker symbol (e.g., AAPL)" & vbCrLf & _
           "2. Enter expiration date (optional)" & vbCrLf & _
           "3. Select CALL or PUT" & vbCrLf & _
           "4. Run 'FetchOptionsData' to load prices" & vbCrLf & vbCrLf & _
           "P&L CALCULATION:" & vbCrLf & _
           "5. Enter 'Stock @ Expiry' price" & vbCrLf & _
           "6. Enter Position (+/- contracts)" & vbCrLf & _
           "7. Adjust Entry Price if needed" & vbCrLf & _
           "8. P&L calculates automatically", vbInformation
End Sub

Public Sub ClearPositions()
    ' Clear all positions and entry prices
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, STRIKE_COL).End(xlUp).Row

    If lastRow >= DATA_START_ROW Then
        ws.Range(ws.Cells(DATA_START_ROW, POSITION_COL), ws.Cells(lastRow, POSITION_COL)).Value = 0
        ws.Range(ws.Cells(DATA_START_ROW, ENTRY_PRICE_COL), ws.Cells(lastRow, ENTRY_PRICE_COL)).ClearContents
    End If

    MsgBox "All positions cleared.", vbInformation
End Sub
