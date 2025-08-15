Option Explicit
Private Const SHEET_SETUP = "Setup"
Private Const SHEET_DATA = "FILE TONG HOA PHU - K HOME"
Private Const SHEET_TIENDO = "TIEN_DO_TT"
Private Const MAX_PAYMENT_PERIODS As Integer = 15

'==== LÀM TRÒN 0 CHỮ SỐ – AWAY FROM ZERO (giống Excel.ROUND) ====
Public Function Round0(ByVal x As Double) As Currency
    If x >= 0 Then
        Round0 = CCur(Int(x + 0.5))
    Else
        Round0 = CCur(-Int(-x + 0.5))
    End If
End Function

Sub TinhToanTongHop_NhaTret_Final()
    Dim wsData As Worksheet, wsTienDo As Worksheet
    Dim r As Range, rowIndex As Long
    Dim nhaTret As clsNhaTret, config As Object
    Dim skippedRows As String, processedCount As Long
    skippedRows = "": processedCount = 0

    On Error GoTo InitializeError
    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)
    Set wsTienDo = ThisWorkbook.Sheets(SHEET_TIENDO)
    Set config = ReadConfig(SHEET_SETUP)
    If config Is Nothing Then Exit Sub
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ProcessError

    Dim selectedRange As Range
    Set selectedRange = Intersect(Selection, wsData.UsedRange)

    For Each r In selectedRange.Rows
        If Not r.EntireRow.Hidden Then
            rowIndex = r.Row

            Dim why As String
            If Not ValidateInputs(wsData, config, rowIndex, why, False) Then
                skippedRows = skippedRows & why & vbCrLf
                GoTo NextRow
            End If

            Set nhaTret = New clsNhaTret
            LoadNhaTretData nhaTret, wsData, config, rowIndex

            ' Tổng % trong TIEN_DO_TT (E, G, I, ...)
            Dim totalPct As Double
            totalPct = SumSchedulePercentages(wsTienDo, nhaTret.TenTienDo)
            nhaTret.XacDinhGiaTriGoc totalPct

            ' Đọc ĐỢT 1 (AB) & ĐỢT 2 (AD) thủ công
            Dim manualDot1 As Currency
            manualDot1 = wsData.Range(config("colStartTienTT") & rowIndex).Value
            If Not IsNumeric(manualDot1) Then manualDot1 = 0
            nhaTret.SetManualDot1Value manualDot1

            Dim manualDot2 As Currency
            manualDot2 = wsData.Range(config("colTienDot2_ThuCong") & rowIndex).Value
            If Not IsNumeric(manualDot2) Then manualDot2 = 0
            nhaTret.SetManualDot2Value manualDot2

            ' Nếu cột T đã có SỐ TIỀN CỌC thì dùng override
            If config.Exists("colPctTienCoc_Data") And Len(config("colPctTienCoc_Data")) > 0 Then
                Dim cocOverride As Variant
                cocOverride = wsData.Range(config("colPctTienCoc_Data") & rowIndex).Value
                If IsNumeric(cocOverride) And cocOverride > 0 Then
                    nhaTret.TienCoc_Override = Round0(CDbl(cocOverride))
                End If
            End If

            nhaTret.TinhTienDoThanhToan
            nhaTret.TaoSoHopDong
            WriteResultsToSheet nhaTret, wsData, config
            WriteValidationTooltips nhaTret, wsData, config

            processedCount = processedCount + 1
        End If
NextRow:
    Next r

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ShowSummaryMsg processedCount, skippedRows
    Exit Sub

InitializeError:
    MsgBox "Lỗi khởi tạo: " & Err.Description, vbCritical, "Lỗi Hệ thống"
    GoTo CleanUp

ProcessError:
    skippedRows = skippedRows & "Dòng " & rowIndex & ": Lỗi xử lý - " & Err.Description & vbCrLf
    Resume Next
End Sub

'----------------- CONFIG -----------------
Private Function ReadConfig(ByVal setupSheetName As String) As Object
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(setupSheetName)
    Set ReadConfig = CreateObject("Scripting.Dictionary")

    With ReadConfig
        .Add "colThanhTienDat_Input", Trim(ws.Range("B1").Value)   ' Q
        '.Add "colThanhTienNha_Input", Trim(ws.Range("B2").Value)   ' R (không dùng)
        .Add "colThanhTien", Trim(ws.Range("B3").Value)            ' S
        .Add "colTenTienDo", Trim(ws.Range("B4").Value)            ' W
        .Add "colStartTienTT", Trim(ws.Range("B5").Value)          ' AB
        .Add "colNgayTT1", Trim(ws.Range("B6").Value)              ' AA
        .Add "colBC_ThanhTien", Trim(ws.Range("B7").Value)         ' BE
        .Add "colBC_TienCoc", Trim(ws.Range("B8").Value)           ' BF (bằng chữ SỐ TIỀN CỌC)
        .Add "colStartBC", Trim(ws.Range("B9").Value)              ' BG (BC_Đợt_1)
        .Add "colPctTienCoc_Data", Trim(ws.Range("B10").Value)     ' T  (SỐ TIỀN CỌC)
        .Add "colLoO", Trim(ws.Range("B11").Value)                 ' F
        .Add "colNgayKy", Trim(ws.Range("B12").Value)              ' H
        .Add "colSoHD", Trim(ws.Range("B13").Value)                ' I
        .Add "colTiLeTT_Dot1_Output", Trim(ws.Range("B14").Value)  ' V
        .Add "colKiemTra", Trim(ws.Range("B15").Value)             ' U
        .Add "colBC_ThanhTien_Dat", Trim(ws.Range("B16").Value)    ' BV
        '.Add "colBC_ThanhTien_Nha", Trim(ws.Range("B17").Value)    ' BW
        .Add "colTienDot2_ThuCong", Trim(ws.Range("B18").Value)    ' AD
        .Add "colLoaiHopDong", Trim(ws.Range("B19").Value)         ' G (NOXH/NOTM)
    End With
    Exit Function
ErrorHandler:
    MsgBox "Lỗi đọc cấu hình sheet '" & setupSheetName & "': " & Err.Description, vbCritical
    Set ReadConfig = Nothing
End Function

'----------------- LOAD INPUTS -----------------
Private Sub LoadNhaTretData(ByVal nhaTret As clsNhaTret, ByVal ws As Worksheet, ByVal conf As Object, ByVal r As Long)
    With nhaTret
        .RowNum = r
        .TongThanhTien = ws.Range(conf("colThanhTien") & r).Value
        .MaSoLo = ws.Range(conf("colLoO") & r).Value
        .NgayKy = ws.Range(conf("colNgayKy") & r).Value
        .TenTienDo = Trim(CStr(ws.Range(conf("colTenTienDo") & r).Value))
        .NgayTTDot1 = ws.Range(conf("colNgayTT1") & r).Value
        .ThanhTienDat_Input = ws.Range(conf("colThanhTienDat_Input") & r).Value
        If conf.Exists("colLoaiHopDong") And Len(conf("colLoaiHopDong")) > 0 Then
            On Error Resume Next
            .LoaiHopDong = Trim$(CStr(ws.Range(conf("colLoaiHopDong") & r).Value)) ' NOXH/NOTM
            On Error GoTo 0
        End If

        ' Đọc SỐ TIỀN CỌC override (nếu người dùng đã nhập Data!T)
        If conf.Exists("colPctTienCoc_Data") And Len(conf("colPctTienCoc_Data")) > 0 Then
            Dim cocVal As Variant
            cocVal = ws.Range(conf("colPctTienCoc_Data") & r).Value
            If IsNumeric(cocVal) And cocVal > 0 Then .TienCoc_Override = Round0(CDbl(cocVal))
        End If
    End With
End Sub

'----------------- SUM TOTAL PCT -----------------
Private Function SumSchedulePercentages(ByVal wsTienDo As Worksheet, ByVal scheduleName As String) As Double
    ' Tổng % ở dạng thập phân (0.3 = 30%)
    Dim total As Double: total = 0
    If Len(Trim(scheduleName)) = 0 Then Exit Function

    Dim r As Long, i As Integer, lastRow As Long
    lastRow = wsTienDo.Cells(wsTienDo.Rows.Count, "C").End(xlUp).Row
    For r = 1 To lastRow
        If Trim$(UCase$(wsTienDo.Cells(r, "C").Value)) = Trim$(UCase$(scheduleName)) Then
            For i = 5 To (5 + (MAX_PAYMENT_PERIODS - 1) * 2) Step 2
                Dim val As Variant: val = wsTienDo.Cells(r, i).Value
                If IsNumeric(val) And Len(CStr(val)) > 0 Then total = total + CDbl(val)
            Next i
            SumSchedulePercentages = total
            Exit Function
        End If
    Next r
End Function

'----------------- WRITE OUTPUTS -----------------
Private Sub WriteResultsToSheet(ByVal nhaTret As clsNhaTret, ByVal ws As Worksheet, ByVal conf As Object)
    Dim r As Long: r = nhaTret.RowNum
    With ws
        .Range(conf("colSoHD") & r).Value = nhaTret.SoHopDong
        .Range(conf("colThanhTien") & r).Value = nhaTret.TongThanhTien

        Dim colTien As Long, colNgay As Long, colBC As Long
        colTien = .Range(conf("colStartTienTT") & 1).Column
        colNgay = .Range(conf("colNgayTT1") & 1).Column
        colBC = .Range(conf("colStartBC") & 1).Column

        Dim bcThanhTienCol As Long, bcThanhTienDatCol As Long
        bcThanhTienCol = .Range(conf("colBC_ThanhTien") & 1).Column
        bcThanhTienDatCol = .Range(conf("colBC_ThanhTien_Dat") & 1).Column

        Dim i As Integer, tgtCol As Long
        For i = 1 To MAX_PAYMENT_PERIODS
            .Cells(r, colTien + (i - 1) * 2).ClearContents
            .Cells(r, colNgay + (i - 1) * 2).ClearContents
            tgtCol = colBC + i - 1
            If tgtCol <> bcThanhTienCol And tgtCol <> bcThanhTienDatCol Then
               .Cells(r, tgtCol).ClearContents
            End If
        Next i

        .Range(conf("colTiLeTT_Dot1_Output") & r).Value = nhaTret.TiLeThanhToanDot1

        ' TIỀN CỌC (Data!T) + bằng chữ ở BF (chỉ KHÔNG HĐMB)
        If Not nhaTret.IsHDMBContract Then
            Dim tienCoc As Currency
            tienCoc = nhaTret.DepositTargetAmount()
            If conf.Exists("colPctTienCoc_Data") And Len(conf("colPctTienCoc_Data")) > 0 Then
                With .Range(conf("colPctTienCoc_Data") & r)
                    .Value = tienCoc
                    On Error Resume Next
                    .NumberFormat = "#,##0"
                    On Error GoTo 0
                End With
            End If
            If tienCoc > 0 Then
                .Range(conf("colBC_TienCoc") & r).Value = vnd(tienCoc)
            Else
                .Range(conf("colBC_TienCoc") & r).ClearContents
            End If
        Else
            If conf.Exists("colPctTienCoc_Data") Then .Range(conf("colPctTienCoc_Data") & r).ClearContents
            .Range(conf("colBC_TienCoc") & r).ClearContents
        End If

        Dim scheduleArray As Variant
        scheduleArray = nhaTret.TienDoThanhToan

        Dim sumKiemTra As Currency: sumKiemTra = 0
        If IsArray(scheduleArray) Then
            On Error Resume Next
            Dim ub As Long: ub = UBound(scheduleArray, 1)
            On Error GoTo 0
            If ub > 0 Then
                For i = 1 To ub
                    Dim soTien As Currency: soTien = scheduleArray(i, 1)
                    .Cells(r, colTien + (i - 1) * 2).Value = soTien
                    tgtCol = colBC + i - 1
                    If tgtCol <> bcThanhTienCol And tgtCol <> bcThanhTienDatCol Then
                        .Cells(r, tgtCol).Value = vnd(soTien)
                    End If
                    If IsDate(scheduleArray(i, 2)) Then
                        .Cells(r, colNgay + (i - 1) * 2).Value = CDate(scheduleArray(i, 2))
                    End If
                    If IsNumeric(soTien) Then sumKiemTra = sumKiemTra + soTien
                Next i
            End If
        End If

        Dim colKiemTra As Long
        colKiemTra = .Range(conf("colKiemTra") & 1).Column
        .Cells(r, colKiemTra).Value = sumKiemTra

        If IsNumeric(nhaTret.TongThanhTien) And nhaTret.TongThanhTien > 0 Then
            .Range(conf("colBC_ThanhTien") & r).Value = vnd(nhaTret.TongThanhTien)
        Else
            .Range(conf("colBC_ThanhTien") & r).ClearContents
        End If

        If IsNumeric(nhaTret.ThanhTienDat_Input) And nhaTret.ThanhTienDat_Input > 0 Then
            .Cells(r, bcThanhTienDatCol).Value = vnd(nhaTret.ThanhTienDat_Input)
        Else
            .Cells(r, bcThanhTienDatCol).ClearContents
        End If
    End With
End Sub

'----------------- UI SUMMARY -----------------
Private Sub ShowSummaryMsg(ByVal processedCount As Long, ByVal skippedRows As String)
    If Len(skippedRows) = 0 Then Exit Sub
    Dim finalMsg As String
    finalMsg = SZ("C19") & vbCrLf & vbCrLf _
             & SZ("C20") & " " & CStr(processedCount) & vbCrLf & vbCrLf _
             & SZ("C21") & vbCrLf & skippedRows
    MsgBoxUni finalMsg, vbExclamation, SZ("C18")
End Sub

Private Function ValidateInputs(ByVal ws As Worksheet, ByVal conf As Object, ByVal r As Long, _
                                ByRef errMsg As String, _
                                Optional ByVal showPopupPerRow As Boolean = False) As Boolean
    Dim ok As Boolean: ok = True
    Dim problems As String

    Dim vNgayKy As Variant:   vNgayKy = ws.Range(conf("colNgayKy") & r).Value
    Dim vTienDo As String:    vTienDo = Trim$(CStr(ws.Range(conf("colTenTienDo") & r).Value))
    Dim vNgayDot1 As Variant: vNgayDot1 = ws.Range(conf("colNgayTT1") & r).Value

    If Not IsDate(vNgayKy) Then problems = problems & Bullet() & SZ("C6") & vbCrLf: ok = False
    If Len(vTienDo) = 0 Then problems = problems & Bullet() & SZ("C4") & vbCrLf: ok = False
    If Not IsDate(vNgayDot1) Then problems = problems & Bullet() & SZ("C12") & vbCrLf: ok = False

    ValidateInputs = ok
    If Not ok Then
        errMsg = SZ("C17") & " " & r & ":" & vbCrLf & problems
        If showPopupPerRow Then MsgBoxUni errMsg, vbExclamation, SZ("C1")
    Else
        errMsg = ""
    End If
End Function

Private Function Bullet() As String
    Bullet = ChrW(8226) & " "
End Function

Public Function SZ(ByVal addr As String) As String
    SZ = ThisWorkbook.Sheets("Setup").Range(addr).Value2
End Function

Public Function MsgBoxUni(ByVal prompt As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "") As VbMsgBoxResult
    MsgBoxUni = MessageBoxW(0, StrPtr(prompt), StrPtr(title), buttons)
End Function

'----------------- Validation tooltips -----------------
Private Sub WriteValidationTooltips(ByVal nhaTret As clsNhaTret, ByVal ws As Worksheet, ByVal conf As Object)
    Dim r As Long: r = nhaTret.RowNum
    Dim scheduleArray As Variant: scheduleArray = nhaTret.TienDoThanhToan
    Dim tongTien As Currency: tongTien = nhaTret.GiaTriGocDeTinhTienDo

    Dim colNgayStart As Long: colNgayStart = ws.Range(conf("colNgayTT1") & 1).Column ' AA
    Dim colTienStart As Long: colTienStart = ws.Range(conf("colStartTienTT") & 1).Column ' AB

    Dim i As Integer
    For i = 0 To (MAX_PAYMENT_PERIODS - 1) * 2
        On Error Resume Next
        ws.Cells(r, colTienStart + i).Validation.Delete
        On Error GoTo 0
    Next i

    If Not IsArray(scheduleArray) Then Exit Sub
    Dim ub As Long
    On Error Resume Next: ub = UBound(scheduleArray, 1): On Error GoTo 0
    If ub = 0 Then Exit Sub

    For i = 1 To ub
        Dim soTien As Variant, ngay As Variant, pct_raw As Variant, days_raw As Variant
        soTien = scheduleArray(i, 1)
        ngay = scheduleArray(i, 2)
        pct_raw = scheduleArray(i, 3)
        days_raw = scheduleArray(i, 4)

        Dim tooltipPct As String
        If IsNumeric(pct_raw) And CDbl(pct_raw) > 0 Then
            tooltipPct = Format(CDbl(pct_raw) * 100, "0.##") & "%"
        ElseIf IsNumeric(soTien) And tongTien > 0 Then
            tooltipPct = Format(soTien / tongTien * 100, "0.##") & "% (*) nhập tay"
        End If
        If Len(tooltipPct) > 0 Then
            With ws.Cells(r, colTienStart + (i - 1) * 2).Validation
                .Add Type:=xlValidateInputOnly
                .InputMessage = tooltipPct
            End With
        End If

        If i > 1 And IsNumeric(days_raw) And days_raw <> "" Then
            With ws.Cells(r, colNgayStart + (i - 1) * 2).Validation
                .Add Type:=xlValidateInputOnly
                .InputMessage = "+" & days_raw & " ngày"
            End With
        End If
    Next i
End Sub
