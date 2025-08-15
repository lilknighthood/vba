Option Explicit
Private Const MAX_PAYMENT_PERIODS As Integer = 15
Private Const SHEET_SETUP As String = "Setup"
Private Const SHEET_TIENDO As String = "TIEN_DO_TT"

'==== ƯU TIÊN DỒN CHÊNH ====
' Mặc định toàn cục: Setup!B20
' Override theo tiến độ: Setup!J1 (từ khoá khớp mờ tên tiến độ), Setup!K1 (quy tắc dồn)
Private Const BALANCE_PREF_GLOBAL_ADDR As String = "B20"
Private Const BALANCE_PREF_KEY_ADDR As String = "J1"
Private Const BALANCE_PREF_VAL_ADDR As String = "K1"

'--- ĐẦU VÀO ---
Public RowNum As Long
Public TongThanhTien As Currency
Public MaSoLo As String
Public NgayKy As Date
Public TenTienDo As String
Public NgayTTDot1 As Date
Public ThanhTienDat_Input As Currency
'Public ThanhTienNha_Input As Currency
Public m_ManualDot1 As Currency
Public m_ManualDot2 As Currency
Public colTienDot2_ThuCong As Currency
Public LoaiHopDong As String        ' NOXH/NOTM (Data!G)
Public TienCoc_Override As Currency ' Nếu Data!T đã nhập tiền cọc

'--- ĐẦU RA ---
Public SoHopDong As String
Public GiaTriGocDeTinhTienDo As Currency
Public GiaTriDeGhiVaoCotCoc As Currency
Public TiLeThanhToanDot1 As Double

'--- PRIVATE ---
Private m_TienDoThanhToan() As Variant
Private m_wsSetup As Worksheet
Private m_wsTienDo As Worksheet
Private m_totalPct As Double

' Round 0 chữ số – dùng Round0 của MainModule
Private Function R0(ByVal x As Double) As Currency
    R0 = Round0(x)
End Function

'========================================================================
' INIT
'========================================================================
Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    Set m_wsSetup = ThisWorkbook.Sheets(SHEET_SETUP)
    Set m_wsTienDo = ThisWorkbook.Sheets(SHEET_TIENDO)
    Exit Sub
ErrorHandler:
    MsgBox "LỖI KHỞI TẠO CLASS!" & vbCrLf & Err.Description, vbCritical
    End
End Sub

'========================================================================
' PROPERTIES
'========================================================================
Public Property Get TienDoThanhToan() As Variant
    TienDoThanhToan = m_TienDoThanhToan
End Property

Public Property Get IsHDMBContract() As Boolean
    Dim nm As String: nm = UCase$(Me.TenTienDo)
    IsHDMBContract = (InStr(1, nm, "HDMB") > 0) _
                  Or (InStr(1, nm, "H" & ChrW(272) & "MB") > 0) _
                  Or (InStr(1, nm, "H" & ChrW(208) & "MB") > 0)
End Property

Public Property Get totalPct() As Double
    totalPct = m_totalPct
End Property
Public Property Let totalPct(ByVal value As Double)
    m_totalPct = value
End Property

'========================================================================
' CORE
'========================================================================
Public Sub XacDinhGiaTriGoc(ByVal totalPercent As Double)
    If Me.ThanhTienDat_Input > 0 Then
        Me.TongThanhTien = Me.ThanhTienDat_Input
    End If
    Me.GiaTriGocDeTinhTienDo = Me.TongThanhTien
    Me.totalPct = totalPercent
End Sub

Public Sub TaoSoHopDong()
    If Len(Trim(Me.MaSoLo)) = 0 Or Not IsDate(Me.NgayKy) Or Len(Trim(Me.TenTienDo)) = 0 Then Exit Sub

    If Not Me.IsHDMBContract Then
        Dim loai As String: loai = UCase$(Trim$(Me.LoaiHopDong))
        If loai <> "NOXH" And loai <> "NOTM" Then loai = "NOXH"
        Me.SoHopDong = Trim$(Me.MaSoLo) & "/" & CStr(Year(Me.NgayKy)) & "/H" & ChrW(208) & "/" & loai & " - HP"
        Exit Sub
    End If

    Dim tpl As String: tpl = GetContractTemplate(Me.TenTienDo)
    If Len(tpl) = 0 Then Exit Sub
    tpl = Replace(tpl, "[NAMKY]", CStr(Year(Me.NgayKy)), 1, -1, vbTextCompare)
    tpl = Replace(tpl, "[LO-O]", Me.MaSoLo, 1, -1, vbTextCompare)
    Me.SoHopDong = tpl
End Sub

Public Sub TinhTienDoThanhToan()
    If Me.GiaTriGocDeTinhTienDo <= 0 Then Exit Sub
    Dim rowSch As Long: rowSch = FindScheduleRow()
    If rowSch = 0 Then Exit Sub

    Dim v As Variant: v = m_wsTienDo.Cells(rowSch, 5).Value
    If IsNumeric(v) Then Me.TiLeThanhToanDot1 = CDbl(v) Else Me.TiLeThanhToanDot1 = 0

    Dim lastPeriod As Integer: lastPeriod = FindLastPaymentPeriod(rowSch)
    If lastPeriod = 0 Then Exit Sub

    BuildPaymentSchedule rowSch, lastPeriod
End Sub

' TIỀN CỌC mục tiêu để cân bằng
Public Function DepositTargetAmount() As Currency
    Dim target As Currency
    If Me.IsHDMBContract Then
        target = R0(Me.TongThanhTien)             ' HĐMB: tổng = S
    Else
        If Me.TienCoc_Override > 0 Then
            target = R0(Me.TienCoc_Override)      ' Non-HĐMB: ưu tiên ô Data!T nếu có
        Else
            target = R0(Me.TongThanhTien * Me.totalPct) ' nếu không thì = S × tổng %
        End If
    End If
    DepositTargetAmount = target
End Function

Public Sub SetManualDot1Value(ByVal manualDot1 As Currency)
    m_ManualDot1 = manualDot1
End Sub
Public Sub SetManualDot2Value(ByVal manualDot2 As Currency)
    m_ManualDot2 = manualDot2
End Sub

'========================================================================
' HELPERS
'========================================================================
Private Function GetContractTemplate(ByVal lookupValue As String) As String
    Dim rng As Range, defTpl As String, tpl As String, i As Long, key As String
    On Error Resume Next
    Set rng = m_wsSetup.Range("G2:H" & m_wsSetup.Cells(m_wsSetup.Rows.Count, "G").End(xlUp).Row)
    On Error GoTo 0
    If rng Is Nothing Or rng.Rows.Count < 1 Then Exit Function

    defTpl = rng.Cells(rng.Rows.Count, 2).Value
    tpl = defTpl

    lookupValue = UCase$(lookupValue)
    For i = 1 To rng.Rows.Count - 1
        key = Trim$(UCase$(rng.Cells(i, 1).Value))
        If Len(key) > 0 And (lookupValue Like "*" & key & "*") Then
            tpl = rng.Cells(i, 2).Value
            Exit For
        End If
    Next i
    GetContractTemplate = tpl
End Function

'--- Chuẩn hoá & kiểm tra ---
Private Function NormalizeVN(ByVal s As String) As String
    s = UCase$(Trim$(CStr(s)))
    s = Replace(s, ChrW(272), "D") ' Đ
    s = Replace(s, ChrW(208), "D") ' Ð
    s = Replace(s, "–", "-"): s = Replace(s, "—", "-"): s = Replace(s, "−", "-")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormalizeVN = s
End Function

Private Function IsValidPref(ByVal v As Variant) As Boolean
    Dim s As String: s = UCase$(Trim$(CStr(v)))
    If s = "LAST" Or s = "FIRST" Or s = "LARGEST" Then IsValidPref = True: Exit Function
    If Len(s) = 0 Then Exit Function
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not (ch Like "[0-9]" Or ch = ",") Then Exit Function
    Next i
    IsValidPref = True
End Function

'---- B20 làm mặc định (fallback); nếu không hợp lệ → LAST ----
Private Function GetBalancePrefGlobal() As String
    Dim v As Variant
    On Error Resume Next
    v = m_wsSetup.Range(BALANCE_PREF_GLOBAL_ADDR).Value
    On Error GoTo 0
    If IsValidPref(v) Then
        GetBalancePrefGlobal = UCase$(Trim$(CStr(v)))
    Else
        GetBalancePrefGlobal = "LAST"
    End If
End Function

'---- J1/K1 override (chỉ khi cả hai có giá trị và J1 khớp mờ tên tiến độ) ----
Private Function GetBalancePref_Override_J1K1(ByVal scheduleName As String) As String
    Dim key As String, pref As Variant
    key = NormalizeVN(m_wsSetup.Range(BALANCE_PREF_KEY_ADDR).Value)
    pref = m_wsSetup.Range(BALANCE_PREF_VAL_ADDR).Value

    If Len(key) = 0 Or Len(Trim$(CStr(pref))) = 0 Then Exit Function

    If NormalizeVN(scheduleName) Like "*" & key & "*" Then
        If IsValidPref(pref) Then
            GetBalancePref_Override_J1K1 = UCase$(Trim$(CStr(pref)))
        End If
    End If
End Function

'---- Quyết định quy tắc dồn: ƯU TIÊN J1/K1, nếu không có thì dùng B20 ----
Private Function ResolveBalancePrefForCurrent() As String
    Dim pref As String
    pref = GetBalancePref_Override_J1K1(Me.TenTienDo)  ' override nếu có
    If pref = "" Then
        pref = GetBalancePrefGlobal()                   ' fallback luôn có
    End If
    ResolveBalancePrefForCurrent = pref
End Function

'---- Thứ tự dồn chênh (áp dụng cho cả diff>0 và diff<0) ----
Private Function BuildBalanceOrder(ByVal lastPeriod As Integer, ByVal pref As String, ByVal diff As Currency) As Variant
    Dim used() As Boolean: ReDim used(1 To lastPeriod)
    Dim list() As Long: ReDim list(1 To lastPeriod)
    Dim count As Long: count = 0
    Dim i As Long, effectivePref As String

    effectivePref = UCase$(Trim$(pref))
    If Len(effectivePref) = 0 Then effectivePref = "LAST"

    Select Case effectivePref
        Case "FIRST"
            For i = 1 To lastPeriod
                count = count + 1: list(count) = i: used(i) = True
            Next i

        Case "LARGEST"
            Dim maxIdx As Long, maxVal As Currency
            maxIdx = 0: maxVal = -1
            For i = 1 To lastPeriod
                If IsNumeric(m_TienDoThanhToan(i, 1)) Then
                    If m_TienDoThanhToan(i, 1) > maxVal Then
                        maxVal = m_TienDoThanhToan(i, 1)
                        maxIdx = i
                    End If
                End If
            Next i
            If maxIdx < 1 Then maxIdx = lastPeriod
            count = count + 1: list(count) = maxIdx: used(maxIdx) = True
            For i = lastPeriod To 1 Step -1
                If Not used(i) Then count = count + 1: list(count) = i: used(i) = True
            Next i

        Case "LAST"
            For i = lastPeriod To 1 Step -1
                count = count + 1: list(count) = i: used(i) = True
            Next i

        Case Else
            Dim parts() As String, s As Variant, n As Long
            parts = Split(effectivePref, ",")
            For Each s In parts
                If Len(Trim$(CStr(s))) > 0 And IsNumeric(s) Then
                    n = CLng(Val(s))
                    If n >= 1 And n <= lastPeriod Then
                        If Not used(n) Then count = count + 1: list(count) = n: used(n) = True
                    End If
                End If
            Next s
            ' Phần còn lại theo LAST
            For i = lastPeriod To 1 Step -1
                If Not used(i) Then count = count + 1: list(count) = i: used(i) = True
            Next i
    End Select

    ReDim Preserve list(1 To count)
    BuildBalanceOrder = list
End Function

Private Function FindScheduleRow() As Long
    Dim i As Long, lastRow As Long
    lastRow = m_wsTienDo.Cells(m_wsTienDo.Rows.Count, "C").End(xlUp).Row
    For i = 1 To lastRow
        If Trim$(UCase$(m_wsTienDo.Cells(i, "C").Value)) = Trim$(UCase$(Me.TenTienDo)) Then
            FindScheduleRow = i
            Exit Function
        End If
    Next i
End Function

Private Function FindLastPaymentPeriod(ByVal scheduleRow As Long) As Integer
    Dim i As Integer, col As Integer, val As Variant
    For i = MAX_PAYMENT_PERIODS To 1 Step -1
        col = (i - 1) * 2 + 5
        val = m_wsTienDo.Cells(scheduleRow, col).Value
        If IsNumeric(val) And Len(CStr(val)) > 0 Then
            FindLastPaymentPeriod = i
            Exit Function
        End If
    Next i
End Function

Private Sub BuildPaymentSchedule(ByVal scheduleRow As Long, ByVal lastPeriod As Integer)
    ' Mảng: [1]=Số tiền, [2]=Ngày TT, [3]=% trên Tổng, [4]=Số ngày cộng
    ReDim m_TienDoThanhToan(1 To lastPeriod, 1 To 4)

    Dim totalPaid As Currency, currentDate As Date
    totalPaid = 0
    currentDate = Me.NgayTTDot1

    Dim pctDot1 As Variant, pctDot2 As Variant
    pctDot1 = m_wsTienDo.Cells(scheduleRow, 5).Value   ' E – Đợt 1
    pctDot2 = m_wsTienDo.Cells(scheduleRow, 7).Value   ' G – Đợt 2

    Dim isDot1Empty As Boolean, isDot2Empty As Boolean, isDot1_25pct As Boolean
    isDot1Empty = Not (IsNumeric(pctDot1) And Len(CStr(pctDot1)) > 0)
    isDot2Empty = Not (IsNumeric(pctDot2) And Len(CStr(pctDot2)) > 0)
    isDot1_25pct = (IsNumeric(pctDot1) And Abs(CDbl(pctDot1) - 0.25) < 0.0001)

    Dim i As Integer
    For i = 1 To lastPeriod
        Dim pay As Currency, pct_raw As Variant, days_raw As Variant
        pct_raw = m_wsTienDo.Cells(scheduleRow, (i - 1) * 2 + 5).Value
        days_raw = m_wsTienDo.Cells(scheduleRow, (i - 2) * 2 + 6).Value

        If Me.IsHDMBContract Then
            ' ==== Các case đặc biệt HĐMB – TRÒN 1 LẦN Ở CUỐI ====
            If i = 1 And isDot1Empty Then
                pay = R0(m_ManualDot1)

            ElseIf i = 2 And isDot1Empty Then
                pay = R0(Me.TongThanhTien * 0.7 - m_ManualDot1)
                If pay < 0 Then pay = 0

            ElseIf i = 2 And isDot1_25pct And isDot2Empty Then
                pay = R0(m_ManualDot2)

            ElseIf i = 3 And isDot1_25pct And isDot2Empty Then
                pay = R0(Me.TongThanhTien * 0.7 - Me.TongThanhTien * 0.25 - m_ManualDot2)
                If pay < 0 Then pay = 0

            Else
                If IsNumeric(pct_raw) Then
                    pay = R0(Me.TongThanhTien * CDbl(pct_raw))
                Else
                    pay = 0
                End If
            End If

        Else
            ' ==== KHÔNG HĐMB ====
            If IsNumeric(pct_raw) Then
                pay = R0(Me.TongThanhTien * CDbl(pct_raw))
            Else
                pay = 0
            End If
        End If

        totalPaid = totalPaid + pay
        m_TienDoThanhToan(i, 1) = pay

        ' Ngày thanh toán
        If i = 1 Then
            m_TienDoThanhToan(i, 2) = currentDate
        ElseIf IsNumeric(days_raw) And Len(CStr(days_raw)) > 0 Then
            currentDate = DateAdd("d", CLng(days_raw), currentDate)
            m_TienDoThanhToan(i, 2) = currentDate
        Else
            m_TienDoThanhToan(i, 2) = ""
        End If

        ' % tooltip
        If Me.TongThanhTien > 0 Then
            m_TienDoThanhToan(i, 3) = Round(CDbl(pay) / CDbl(Me.TongThanhTien), 6)
        Else
            m_TienDoThanhToan(i, 3) = 0
        End If

        m_TienDoThanhToan(i, 4) = days_raw
    Next i

    ' ====== CÂN BẰNG PHẦN DƯ ======
    Dim targetTotal As Currency: targetTotal = DepositTargetAmount()
    Dim diff As Currency: diff = targetTotal - totalPaid

    If diff <> 0 Then
        Dim pref As String: pref = ResolveBalancePrefForCurrent()
        Dim order As Variant: order = BuildBalanceOrder(lastPeriod, pref, diff)

        Dim j As Long, idx As Variant, adj As Currency, canReduce As Currency
        For Each idx In order
            j = CLng(idx)
            If j >= 1 And j <= lastPeriod Then
                If m_TienDoThanhToan(j, 1) > 0 Then
                    If diff > 0 Then
                        adj = diff
                    Else
                        canReduce = m_TienDoThanhToan(j, 1)
                        If canReduce + diff >= 0 Then
                            adj = diff
                        Else
                            adj = -canReduce
                        End If
                    End If
                    m_TienDoThanhToan(j, 1) = m_TienDoThanhToan(j, 1) + adj
                    diff = diff - adj
                    If diff = 0 Then Exit For
                End If
            End If
        Next idx
    End If

    ' Cập nhật lại % tooltip
    Dim i2 As Integer
    For i2 = 1 To lastPeriod
        Dim pay2 As Currency: pay2 = m_TienDoThanhToan(i2, 1)
        If Me.TongThanhTien > 0 Then
            m_TienDoThanhToan(i2, 3) = Round(CDbl(pay2) / CDbl(Me.TongThanhTien), 6)
        Else
            m_TienDoThanhToan(i2, 3) = 0
        End If
    Next i2
End Sub
