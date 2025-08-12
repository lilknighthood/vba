Option Explicit
Private Const MAX_PAYMENT_PERIODS As Integer = 15
Private Const SHEET_SETUP As String = "Setup"
Private Const SHEET_TIENDO As String = "TIEN_DO_TT"

'--- Ð?U VÀO ---
Public RowNum As Long
Public TongThanhTien As Currency
Public MaSoLo As String
Public NgayKy As Date
Public TenTienDo As String
Public NgayTTDot1 As Date
Public ThanhTienDat_Input As Currency
Public ThanhTienNha_Input As Currency
Public m_ManualDot1 As Currency
Public m_ManualDot2 As Currency
Public colTienDot2_ThuCong As Currency
Public LoaiHopDong As String        ' NOXH ho?c NOTM (t? c?t G c?a Data)
Public TienCoc_Override As Currency ' N?u Data!T dã nh?p s?n ti?n c?c, s? uu tiên dùng

'--- Ð?U RA ---
Public SoHopDong As String
Public GiaTriGocDeTinhTienDo As Currency
Public GiaTriDeGhiVaoCotCoc As Currency
Public TiLeThanhToanDot1 As Double

'--- RIÊNG TU ---
Private m_TienDoThanhToan() As Variant
Private m_wsSetup As Worksheet
Private m_wsTienDo As Worksheet
Private m_totalPct As Double   ' T?ng % l?ch (t?ng các % c?t E,G,I,...)

' Round 0 ch? s? – away from zero (dùng l?i hàm Public Round0 ? module)
Private Function R0(ByVal x As Double) As Currency
    R0 = Round0(x)
End Function

'========================================================================
'                       KH?I T?O CLASS
'========================================================================
Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    Set m_wsSetup = ThisWorkbook.Sheets(SHEET_SETUP)
    Set m_wsTienDo = ThisWorkbook.Sheets(SHEET_TIENDO)
    Exit Sub
ErrorHandler:
    MsgBox "L?I KH?I T?O CLASS!" & vbCrLf & Err.Description, vbCritical
    End
End Sub

'========================================================================
'                       THU?C TÍNH
'========================================================================
Public Property Get TienDoThanhToan() As Variant
    TienDoThanhToan = m_TienDoThanhToan
End Property

Public Property Get IsHDMBContract() As Boolean
    Dim nm As String: nm = UCase$(Me.TenTienDo)
    ' B?t c? HÐMB và HÐMB (l?i font Ð)
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
'                       PHUONG TH?C CHÍNH
'========================================================================
Public Sub XacDinhGiaTriGoc(ByVal totalPercent As Double)
    ' THÀNH_TI?N = THÀNH_TI?N_Ð?T + THÀNH_TI?N_NHÀ (n?u 2 c?t này có)
    If Me.ThanhTienDat_Input > 0 Or Me.ThanhTienNha_Input > 0 Then
        Me.TongThanhTien = Me.ThanhTienDat_Input + Me.ThanhTienNha_Input
    End If
    Me.GiaTriGocDeTinhTienDo = Me.TongThanhTien
    Me.totalPct = totalPercent
End Sub

Public Sub TaoSoHopDong()
    If Len(Trim(Me.MaSoLo)) = 0 Or Not IsDate(Me.NgayKy) Or Len(Trim(Me.TenTienDo)) = 0 Then Exit Sub

    ' KHÔNG HÐMB: [LO-O]/[NAMKY]/HÐ/[LO?I_H?P_Ð?NG] - HP
    If Not Me.IsHDMBContract Then
        Dim loai As String: loai = UCase$(Trim$(Me.LoaiHopDong))
        If loai <> "NOXH" And loai <> "NOTM" Then loai = "NOXH"
        Me.SoHopDong = Trim$(Me.MaSoLo) & "/" & CStr(Year(Me.NgayKy)) & "/H" & ChrW(208) & "/" & loai & " - HP"
        Exit Sub
    End If

    ' HÐMB: dùng template G:H trên Setup (nhu tru?c)
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

    ' % Ð?t 1 d? hi?n th?
    Dim v As Variant: v = m_wsTienDo.Cells(rowSch, 5).value
    If IsNumeric(v) Then Me.TiLeThanhToanDot1 = CDbl(v) Else Me.TiLeThanhToanDot1 = 0

    Dim lastPeriod As Integer: lastPeriod = FindLastPaymentPeriod(rowSch)
    If lastPeriod = 0 Then Exit Sub

    BuildPaymentSchedule rowSch, lastPeriod
End Sub

'==> TI?N C?C M?C TIÊU (d? cân b?ng t?ng các d?t) <==
Public Function DepositTargetAmount() As Currency
    Dim target As Currency
    If Me.IsHDMBContract Then
        target = R0(Me.TongThanhTien)  ' HÐMB: t?ng d?t = THÀNH_TI?N
    Else
        If Me.TienCoc_Override > 0 Then
            target = R0(Me.TienCoc_Override)                        ' Uu tiên s? nh?p tay t?i Data!T
        Else
            target = R0(Me.TongThanhTien * Me.totalPct)             ' Tính = S × (t?ng % các d?t)
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
'                       HÀM H? TR? (PRIVATE)
'========================================================================
Private Function GetContractTemplate(ByVal lookupValue As String) As String
    Dim rng As Range, defTpl As String, tpl As String, i As Long
    Dim key As String
    
    On Error Resume Next
    Set rng = m_wsSetup.Range("G2:H" & m_wsSetup.Cells(m_wsSetup.Rows.Count, "G").End(xlUp).row)
    On Error GoTo 0
    
    If rng Is Nothing Or rng.Rows.Count < 1 Then Exit Function
    
    defTpl = rng.Cells(rng.Rows.Count, 2).value
    tpl = defTpl
    
    lookupValue = UCase$(lookupValue)
    
    For i = 1 To rng.Rows.Count - 1
        key = Trim$(UCase$(rng.Cells(i, 1).value))
        If Len(key) > 0 And (lookupValue Like "*" & key & "*") Then
            tpl = rng.Cells(i, 2).value
            Exit For
        End If
    Next i
    
    GetContractTemplate = tpl
End Function

Private Function FindScheduleRow() As Long
    Dim i As Long, lastRow As Long
    lastRow = m_wsTienDo.Cells(m_wsTienDo.Rows.Count, "C").End(xlUp).row
    For i = 1 To lastRow
        If Trim$(UCase$(m_wsTienDo.Cells(i, "C").value)) = Trim$(UCase$(Me.TenTienDo)) Then
            FindScheduleRow = i
            Exit Function
        End If
    Next i
End Function

Private Function FindLastPaymentPeriod(ByVal scheduleRow As Long) As Integer
    Dim i As Integer, col As Integer, val As Variant
    For i = MAX_PAYMENT_PERIODS To 1 Step -1
        col = (i - 1) * 2 + 5
        val = m_wsTienDo.Cells(scheduleRow, col).value
        If IsNumeric(val) And Len(CStr(val)) > 0 Then
            FindLastPaymentPeriod = i
            Exit Function
        End If
    Next i
End Function

Private Sub BuildPaymentSchedule(ByVal scheduleRow As Long, ByVal lastPeriod As Integer)
    ' M?ng: [1]=S? ti?n, [2]=Ngày TT, [3]=% th?c t? trên T?ng, [4]=S? ngày c?ng
    ReDim m_TienDoThanhToan(1 To lastPeriod, 1 To 4)

    Dim totalPaid As Currency, currentDate As Date
    totalPaid = 0
    currentDate = Me.NgayTTDot1

    Dim pctDot1 As Variant, pctDot2 As Variant
    pctDot1 = m_wsTienDo.Cells(scheduleRow, 5).value     ' E – Ð?t 1
    pctDot2 = m_wsTienDo.Cells(scheduleRow, 7).value     ' G – Ð?t 2

    Dim isDot1Empty As Boolean, isDot2Empty As Boolean, isDot1_25pct As Boolean
    isDot1Empty = Not (IsNumeric(pctDot1) And Len(CStr(pctDot1)) > 0)
    isDot2Empty = Not (IsNumeric(pctDot2) And Len(CStr(pctDot2)) > 0)
    isDot1_25pct = (IsNumeric(pctDot1) And Abs(CDbl(pctDot1) - 0.25) < 0.0001)

    Dim i As Integer
    For i = 1 To lastPeriod
        Dim pay As Currency, pct_raw As Variant, days_raw As Variant
        pct_raw = m_wsTienDo.Cells(scheduleRow, (i - 1) * 2 + 5).value
        days_raw = m_wsTienDo.Cells(scheduleRow, (i - 2) * 2 + 6).value

        If Me.IsHDMBContract Then
            ' ==== 3 tru?ng h?p HÐMB (gi? logic cu, nhung dùng R0) ====
            If i = 1 And isDot1Empty Then
                pay = R0(m_ManualDot1)

            ElseIf i = 2 And isDot1Empty Then
                pay = R0(Me.TongThanhTien * 0.7) - R0(m_ManualDot1)
                If pay < 0 Then pay = 0

            ElseIf i = 2 And isDot1_25pct And isDot2Empty Then
                pay = R0(m_ManualDot2)

            ElseIf i = 3 And isDot1_25pct And isDot2Empty Then
                pay = R0(Me.TongThanhTien * 0.7) - R0(Me.TongThanhTien * 0.25) - R0(m_ManualDot2)
                If pay < 0 Then pay = 0

            Else
                If IsNumeric(pct_raw) Then
                    pay = R0(Me.TongThanhTien * CDbl(pct_raw))
                Else
                    pay = 0
                End If
            End If

        Else
            ' ==== KHÔNG HÐMB: tính theo % t?ng d?t (R0) – s? cân b?ng theo TI?N_C?C ====
            If IsNumeric(pct_raw) Then
                pay = R0(Me.TongThanhTien * CDbl(pct_raw))
            Else
                pay = 0
            End If
        End If

        totalPaid = totalPaid + pay
        m_TienDoThanhToan(i, 1) = pay ' S? ti?n

        ' Ngày thanh toán
        If i = 1 Then
            m_TienDoThanhToan(i, 2) = currentDate
        ElseIf IsNumeric(days_raw) And Len(CStr(days_raw)) > 0 Then
            currentDate = DateAdd("d", CLng(days_raw), currentDate)
            m_TienDoThanhToan(i, 2) = currentDate
        Else
            m_TienDoThanhToan(i, 2) = ""
        End If

        ' % th?c nh?n trên T?NG (tooltip)
        If Me.TongThanhTien > 0 Then
            m_TienDoThanhToan(i, 3) = Round(CDbl(pay) / CDbl(Me.TongThanhTien), 6)
        Else
            m_TienDoThanhToan(i, 3) = 0
        End If

        m_TienDoThanhToan(i, 4) = days_raw
    Next i

    ' ====== CÂN B?NG PH?N DU ======
    Dim targetTotal As Currency
    targetTotal = DepositTargetAmount()                ' HÐMB: = S; KHÔNG HÐMB: = ti?n c?c

    Dim diff As Currency
    diff = targetTotal - totalPaid
    If diff <> 0 Then
        Dim j As Long, adj As Currency, canReduce As Currency
        For j = lastPeriod To 1 Step -1
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
        Next j
    End If

    ' C?p nh?t l?i % sau cân b?ng (tooltip)
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


