Option Explicit
Private Const SHEET_SETUP = "Setup"
Private Const SHEET_DATA = "FILE TONG HOA PHU - K HOME"
Private Const SHEET_TIENDO = "TIEN_DO_TT"
Private Const MAX_PAYMENT_PERIODS As Integer = 15

'==== LÀM TRÒN 0 CH? S? – AWAY FROM ZERO (gi?ng Excel.ROUND) ====
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

    ' Khoi tao worksheet va config
    On Error GoTo InitializeError
    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)
    Set wsTienDo = ThisWorkbook.Sheets(SHEET_TIENDO)
    Set config = ReadConfig(SHEET_SETUP)
    If config Is Nothing Then Exit Sub
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ProcessError

    ' Xac dinh vung duoc chon hop le
    Dim selectedRange As Range
    Set selectedRange = Intersect(Selection, wsData.UsedRange)

    ' Vong lap qua tung dong
    For Each r In selectedRange.Rows
        If Not r.EntireRow.Hidden Then
            rowIndex = r.row

            Dim why As String
            If Not ValidateInputs(wsData, config, rowIndex, why, False) Then
                skippedRows = skippedRows & why & vbCrLf
                GoTo NextRow
            End If

            Set nhaTret = New clsNhaTret
            LoadNhaTretData nhaTret, wsData, config, rowIndex

            ' Tinh tong % tien do
            Dim totalPct As Double
            totalPct = SumSchedulePercentages(wsTienDo, nhaTret.TenTienDo)
            nhaTret.XacDinhGiaTriGoc totalPct

            ' Doc tien DOT 1 thu cong (cot AB)
            Dim manualDot1 As Currency
            manualDot1 = wsData.Range(config("colStartTienTT") & rowIndex).value
            If IsNumeric(manualDot1) Then
                nhaTret.SetManualDot1Value manualDot1
            Else
                nhaTret.SetManualDot1Value 0
            End If

            ' Doc tien DOT 2 thu cong (cot AD)
            Dim manualDot2 As Currency
            manualDot2 = wsData.Range(config("colTienDot2_ThuCong") & rowIndex).value
            If IsNumeric(manualDot2) Then
                nhaTret.SetManualDot2Value manualDot2
            Else
                nhaTret.SetManualDot2Value 0
            End If

            ' Tinh tien do, tao hop dong, ghi ket qua
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
    MsgBox "Loi khoi tao: " & Err.Description, vbCritical, "Loi He Thong"
    GoTo CleanUp

ProcessError:
    skippedRows = skippedRows & "Dong " & rowIndex & ": Loi xu ly - " & Err.Description & vbCrLf
    Resume Next
End Sub


'----------------- CONFIG -----------------
Private Function ReadConfig(ByVal setupSheetName As String) As Object
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(setupSheetName)
    Set ReadConfig = CreateObject("Scripting.Dictionary")

    With ReadConfig
        .Add "colThanhTienDat_Input", Trim(ws.Range("B1").value)   ' Q
        .Add "colThanhTienNha_Input", Trim(ws.Range("B2").value)   ' R
        .Add "colThanhTien", Trim(ws.Range("B3").value)            ' S
        .Add "colTenTienDo", Trim(ws.Range("B4").value)            ' T
        .Add "colStartTienTT", Trim(ws.Range("B5").value)          ' AB
        .Add "colNgayTT1", Trim(ws.Range("B6").value)              ' AA
        .Add "colBC_ThanhTien", Trim(ws.Range("B7").value)         ' BE
        .Add "colBC_TienCoc", Trim(ws.Range("B8").value)           ' BF
        .Add "colStartBC", Trim(ws.Range("B9").value)              ' BG (BC_Dat_1)
        .Add "colCoc_NonHDMB_Output", Trim(ws.Range("B10").value)  ' F
        .Add "colLoO", Trim(ws.Range("B11").value)                 ' F
        .Add "colNgayKy", Trim(ws.Range("B12").value)              ' H
        .Add "colSoHD", Trim(ws.Range("B13").value)                ' I
        .Add "colTiLeTT_Dot1_Output", Trim(ws.Range("B14").value)  ' V
        .Add "colKiemTra", Trim(ws.Range("B15").value)             ' U
        .Add "colBC_ThanhTien_Dat", Trim(ws.Range("B16").value)    ' BV
        .Add "colBC_ThanhTien_Nha", Trim(ws.Range("B17").value)    ' BW
        .Add "colTienDot2_ThuCong", Trim(ws.Range("B18").value)    ' AD
    End With
    Exit Function
ErrorHandler:
    MsgBox "Loi doc cau hinh tu sheet '" & setupSheetName & "': " & Err.Description, vbCritical
    Set ReadConfig = Nothing
End Function

'----------------- LOAD INPUTS -----------------
Private Sub LoadNhaTretData(ByVal nhaTret As clsNhaTret, ByVal ws As Worksheet, ByVal conf As Object, ByVal r As Long)
    With nhaTret
        .RowNum = r
        .TongThanhTien = ws.Range(conf("colThanhTien") & r).value
        .MaSoLo = ws.Range(conf("colLoO") & r).value
        .NgayKy = ws.Range(conf("colNgayKy") & r).value
        .TenTienDo = Trim(CStr(ws.Range(conf("colTenTienDo") & r).value))
        .NgayTTDot1 = ws.Range(conf("colNgayTT1") & r).value
        .ThanhTienDat_Input = ws.Range(conf("colThanhTienDat_Input") & r).value
        .ThanhTienNha_Input = ws.Range(conf("colThanhTienNha_Input") & r).value
    End With
End Sub

'----------------- SUM TOTAL PCT -----------------
Private Function SumSchedulePercentages(ByVal wsTienDo As Worksheet, ByVal scheduleName As String) As Double
    ' Tra ve tong % o dang thap phan (0.3 = 30%)
    Dim total As Double: total = 0
    If Len(Trim(scheduleName)) = 0 Then Exit Function
    Dim r As Long, i As Integer, lastRow As Long
    lastRow = wsTienDo.Cells(wsTienDo.Rows.Count, "C").End(xlUp).row
    For r = 1 To lastRow
        If Trim$(UCase$(wsTienDo.Cells(r, "C").value)) = Trim$(UCase$(scheduleName)) Then
            For i = 5 To (5 + (MAX_PAYMENT_PERIODS - 1) * 2) Step 2
                Dim val As Variant: val = wsTienDo.Cells(r, i).value
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
        ' So HD
        .Range(conf("colSoHD") & r).value = nhaTret.SoHopDong
        ' GHI VÀO C?T THÀNH_TI?N (S)
        .Range(conf("colThanhTien") & r).value = nhaTret.TongThanhTien

        ' Cot moc
        Dim colTien As Long, colNgay As Long, colBC As Long
        colTien = .Range(conf("colStartTienTT") & 1).Column
        colNgay = .Range(conf("colNgayTT1") & 1).Column
        colBC = .Range(conf("colStartBC") & 1).Column

        ' Cac cot "bang chu" can BO QUA khi clear
        Dim bcThanhTienCol As Long, bcThanhTienDatCol As Long, bcThanhTienNhaCol As Long
        bcThanhTienCol = .Range(conf("colBC_ThanhTien") & 1).Column      ' BE
        bcThanhTienDatCol = .Range(conf("colBC_ThanhTien_Dat") & 1).Column ' BV
        bcThanhTienNhaCol = .Range(conf("colBC_ThanhTien_Nha") & 1).Column ' BW

        ' CLEAR (bo qua BE/BV/BW)
        Dim i As Integer, tgtCol As Long
        For i = 1 To MAX_PAYMENT_PERIODS
            .Cells(r, colTien + (i - 1) * 2).ClearContents
            .Cells(r, colNgay + (i - 1) * 2).ClearContents
            tgtCol = colBC + i - 1
            If tgtCol <> bcThanhTienCol _
               And tgtCol <> bcThanhTienDatCol _
               And tgtCol <> bcThanhTienNhaCol Then
               .Cells(r, tgtCol).ClearContents
            End If
        Next i

        ' Ti le dot 1
        .Range(conf("colTiLeTT_Dot1_Output") & r).value = nhaTret.TiLeThanhToanDot1

        ' ===== %_TIEN_COC =====
        If Not nhaTret.IsHDMBContract Then
            Dim tienCocVal As Currency
            tienCocVal = Round0(nhaTret.TongThanhTien * nhaTret.totalPct) ' dùng Round0 d?ng b?
            .Range(conf("colCoc_NonHDMB_Output") & r).value = tienCocVal

            If tienCocVal > 0 Then
                .Range(conf("colBC_TienCoc") & r).value = vnd(tienCocVal) ' bang chu coc
            Else
                .Range(conf("colBC_TienCoc") & r).ClearContents
            End If
        Else
            .Range(conf("colCoc_NonHDMB_Output") & r).ClearContents
            .Range(conf("colBC_TienCoc") & r).ClearContents
        End If

        ' Lich thanh toan + BC_Dat_i
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
                    .Cells(r, colTien + (i - 1) * 2).value = soTien
                    tgtCol = colBC + i - 1
                    If tgtCol <> bcThanhTienCol _
                       And tgtCol <> bcThanhTienDatCol _
                       And tgtCol <> bcThanhTienNhaCol Then
                        .Cells(r, tgtCol).value = vnd(soTien) ' bang chu dat
                    End If
                    If IsDate(scheduleArray(i, 2)) Then
                        .Cells(r, colNgay + (i - 1) * 2).value = CDate(scheduleArray(i, 2))
                    End If
                    If IsNumeric(soTien) Then sumKiemTra = sumKiemTra + soTien
                Next i
            End If
        End If

        ' KIEM_TRA (B15)
        Dim colKiemTra As Long
        colKiemTra = .Range(conf("colKiemTra") & 1).Column
        .Cells(r, colKiemTra).value = sumKiemTra
        ' muon bang chu: .Cells(r, colKiemTra).Value = vnd(sumKiemTra)

        ' BC_THANH_TIEN (BE) – ghi CUOI
        If IsNumeric(nhaTret.TongThanhTien) And nhaTret.TongThanhTien > 0 Then
            .Range(conf("colBC_ThanhTien") & r).value = vnd(nhaTret.TongThanhTien)
        Else
            .Range(conf("colBC_ThanhTien") & r).ClearContents
        End If

        ' BC_THANH_TIEN_DAT (BV) – bang chu
        If IsNumeric(nhaTret.ThanhTienDat_Input) And nhaTret.ThanhTienDat_Input > 0 Then
            .Cells(r, bcThanhTienDatCol).value = vnd(nhaTret.ThanhTienDat_Input)
        Else
            .Cells(r, bcThanhTienDatCol).ClearContents
        End If

        ' BC_THANH_TIEN_NHA (BW) – bang chu
        If IsNumeric(nhaTret.ThanhTienNha_Input) And nhaTret.ThanhTienNha_Input > 0 Then
            .Cells(r, bcThanhTienNhaCol).value = vnd(nhaTret.ThanhTienNha_Input)
        Else
            .Cells(r, bcThanhTienNhaCol).ClearContents
        End If
    End With
End Sub

'----------------- UI SUMMARY -----------------
Private Sub ShowSummaryMsg(ByVal processedCount As Long, ByVal skippedRows As String)
    ' Chi hien thong bao neu co dong bi bo qua (loi)
    If Len(skippedRows) = 0 Then Exit Sub
    
    Dim finalMsg As String
    finalMsg = SZ("C19") & vbCrLf & vbCrLf _
             & SZ("C20") & " " & CStr(processedCount) & vbCrLf & vbCrLf _
             & SZ("C21") & vbCrLf & skippedRows
    
    ' Doi sang vbExclamation de ro la canh bao/loi
    MsgBoxUni finalMsg, vbExclamation, SZ("C18")
End Sub


Private Function ValidateInputs(ByVal ws As Worksheet, ByVal conf As Object, ByVal r As Long, _
                                ByRef errMsg As String, _
                                Optional ByVal showPopupPerRow As Boolean = False) As Boolean
    Dim ok As Boolean: ok = True
    Dim problems As String

    Dim vNgayKy As Variant:   vNgayKy = ws.Range(conf("colNgayKy") & r).value
    Dim vTienDo As String:    vTienDo = Trim$(CStr(ws.Range(conf("colTenTienDo") & r).value))
    Dim vNgayDot1 As Variant: vNgayDot1 = ws.Range(conf("colNgayTT1") & r).value

    If Not IsDate(vNgayKy) Then
        problems = problems & Bullet() & SZ("C6") & vbCrLf
        ok = False
    End If
    If Len(vTienDo) = 0 Then
        problems = problems & Bullet() & SZ("C4") & vbCrLf
        ok = False
    End If
    If Not IsDate(vNgayDot1) Then
        problems = problems & Bullet() & SZ("C12") & vbCrLf
        ok = False
    End If

    ValidateInputs = ok
    If Not ok Then
        errMsg = SZ("C17") & " " & r & ":" & vbCrLf & problems
        If showPopupPerRow Then
            MsgBoxUni errMsg, vbExclamation, SZ("C1")
        End If
    Else
        errMsg = ""
    End If
End Function

' TONG KET CUOI CUNG — 1 POPUP Duy NHAT, TOAN BO LOI
'========================================================================
'         HAM DOC VAN BAN TU SHEET SETUP
'========================================================================
' Bullet Unicode (•)
Private Function Bullet() As String
    Bullet = ChrW(8226) & " "
End Function

Public Function SZ(ByVal addr As String) As String
    SZ = ThisWorkbook.Sheets("Setup").Range(addr).Value2  ' doc Unicode tu o
End Function

'========================================================================
'         HÀM MSGBOX H? TR? UNICODE (DÙNG WINDOWS API)
'========================================================================
Public Function MsgBoxUni(ByVal prompt As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "") As VbMsgBoxResult
    ' phiên b?n dùng MessageBoxW dã có s?n trong project c?a b?n
    MsgBoxUni = MessageBoxW(0, StrPtr(prompt), StrPtr(title), buttons)
End Function

'----------------- WRITE VALIDATION TOOLTIPS (v3) -----------------
Private Sub WriteValidationTooltips(ByVal nhaTret As clsNhaTret, ByVal ws As Worksheet, ByVal conf As Object)
    Dim r As Long: r = nhaTret.RowNum
    Dim scheduleArray As Variant: scheduleArray = nhaTret.TienDoThanhToan
    Dim tongTien As Currency: tongTien = nhaTret.GiaTriGocDeTinhTienDo

    Dim colNgayStart As Long: colNgayStart = ws.Range(conf("colNgayTT1") & 1).Column ' Cot AA
    Dim colTienStart As Long: colTienStart = ws.Range(conf("colStartTienTT") & 1).Column ' Cot AB

    ' ==== XOA VALIDATION CU ====
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

    ' ==== GHI TOOLTIP ====
    For i = 1 To ub
        Dim soTien As Variant, ngay As Variant, pct_raw As Variant, days_raw As Variant
        soTien = scheduleArray(i, 1)
        ngay = scheduleArray(i, 2)
        pct_raw = scheduleArray(i, 3)
        days_raw = scheduleArray(i, 4)

        ' -- TIEN: Tooltip phan tram --
        Dim tooltipPct As String
        If IsNumeric(pct_raw) And CDbl(pct_raw) > 0 Then
            tooltipPct = Format(CDbl(pct_raw) * 100, "0.##") & "%"
        ElseIf IsNumeric(soTien) And tongTien > 0 Then
            tooltipPct = Format(soTien / tongTien * 100, "0.##") & "% (*) nhap tay"
        End If
        If Len(tooltipPct) > 0 Then
            With ws.Cells(r, colTienStart + (i - 1) * 2).Validation
                .Add Type:=xlValidateInputOnly
                .InputMessage = tooltipPct
            End With
        End If

        ' -- NGAY: Tooltip so ngay (tu dot 2 tro di) --
        If i > 1 And IsNumeric(days_raw) And days_raw <> "" Then
            With ws.Cells(r, colNgayStart + (i - 1) * 2).Validation
                .Add Type:=xlValidateInputOnly
                .InputMessage = "+" & days_raw & " ngay"
            End With
        End If
    Next i
End Sub


