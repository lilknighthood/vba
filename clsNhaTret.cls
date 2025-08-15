Option Explicit
Private Const MAX_PAYMENT_PERIODS As Integer = 15
Private Const SHEET_SETUP As String = "Setup"
Private Const SHEET_TIENDO As String = "TIEN_DO_TT"

'==== C?U HÌNH D?N CHÊNH ====
' M?c d?nh toàn c?c (áp d?ng khi ti?n d? không có c?u hình riêng):
Private Const BALANCE_PREF_GLOBAL_ADDR As String = "B20"   ' Setup!B20
' B?ng c?u hình theo TI?N_Ð?: Setup!J:K (J = tên ti?n d?, K = quy t?c d?n)
Private Const BALANCE_PREF_KEY_COL As String = "J"
Private Const BALANCE_PREF_VAL_COL As String = "K"

'--- Ð?U VÀO ---
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
Public LoaiHopDong As String        ' NOXH/NOTM (t? c?t G)
Public TienCoc_Override As Currency ' N?u Data!T dã nh?p ti?n c?c

'--- Ð?U RA ---
Public SoHopDong As String
Public GiaTriGocDeTinhTienDo As Currency
Public GiaTriDeGhiVaoCotCoc As Currency
Public TiLeThanhToanDot1 As Double

'--- PRIVATE ---
Private m_TienDoThanhToan() As Variant
Private m_wsSetup As Worksheet
Private m_wsTienDo As Worksheet
Private m_totalPct As Double

' Round 0 ch? s? – dùng Round0 c?a MainModule
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
    MsgBox "L?I KH?I T?O CLASS!" & vbCrLf & Err.Description, vbCritical
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

    Dim v As Variant: v = m_wsTienDo.Cells(rowSch, 5).value
    If IsNumeric(v) Then Me.TiLeThanhToanDot1 = CDbl(v) Else Me.TiLeThanhToanDot1 = 0

    Dim lastPeriod As Integer: lastPeriod = FindLastPaymentPeriod(rowSch)
    If lastPeriod = 0 Then Exit Sub

    BuildPaymentSchedule rowSch, lastPeriod
End Sub

' TI?N C?C m?c tiêu d? cân b?ng
Public Function DepositTargetAmount() As Currency
    Dim target As Currency
    If Me.IsHDMBContract Then
        target = R0(Me.TongThanhTien)
    Else
        If Me.TienCoc_Override > 0 Then
            target = R0(Me.TienCoc_Override)
        Else
            target = R0(Me.TongThanhTien * Me.totalPct)
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
    Set rng = m_wsSetup.Range("G2:H" & m_wsSetup.Cells(m_wsSetup.Rows.count, "G").End(xlUp).Row)
    On Error GoTo 0
    If rng Is Nothing Or rng.Rows.count < 1 Then Exit Function

    defTpl = rng.Cells(rng.Rows.count, 2).value
    tpl = defTpl

    lookupValue = UCase$(lookupValue)
    For i = 1 To rng.Rows.count - 1
        key = Trim$(UCase$(rng.Cells(i, 1).value))
        If Len(key) > 0 And (lookupValue Like "*" & key & "*") Then
            tpl = rng.Cells(i, 2).value
            Exit For
        End If
    Next i
    GetContractTemplate = tpl
End Function

' Uu tiên d?n m?c d?nh toàn c?c
Private Function GetBalancePrefGlobal() As String
    Dim v As Variant
    On Error Resume Next
    v = m_wsSetup.Range(BALANCE_PREF_GLOBAL_ADDR).value
    On Error GoTo 0
    GetBalancePrefGlobal = UCase$(Trim$(CStr(v)))
    If GetBalancePrefGlobal = "" Then GetBalancePrefGlobal = "LAST"
End Function

' Uu tiên d?n theo tên ti?n d? (J:K)
Private Function GetBalancePrefForSchedule(ByVal scheduleName As String) As String
    Dim lastRow As Long, r As Long, key As String, pref As String
    If Len(Trim$(scheduleName)) = 0 Then Exit Function
    lastRow = m_wsSetup.Cells(m_wsSetup.Rows.count, BALANCE_PREF_KEY_COL).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    For r = 2 To lastRow
        key = UCase$(Trim$(CStr(m_wsSetup.Cells(r, BALANCE_PREF_KEY_COL).value)))
        If key = UCase$(Trim$(scheduleName)) Then
            pref = UCase$(Trim$(CStr(m_wsSetup.Cells(r, BALANCE_PREF_VAL_COL).value)))
            If Len(pref) > 0 Then
                GetBalancePrefForSchedule = pref
                Exit Function
            End If
        End If
    Next r
End Function

Private Function ResolveBalancePrefForCurrent() As String
    Dim pref As String
    pref = GetBalancePrefForSchedule(Me.TenTienDo)
    If pref = "" Then pref = GetBalancePrefGlobal()
    ResolveBalancePrefForCurrent = pref
End Function

' Xây th? t? các d?t d? d?n chênh.
' - diff > 0 (thi?u): dùng uu tiên theo ti?n d? (pref).
' - diff < 0 (du): ALWAYS LAST (t? d?t cu?i v? d?u).
Private Function BuildBalanceOrder(ByVal lastPeriod As Integer, ByVal pref As String, ByVal diff As Currency) As Variant
    Dim used() As Boolean: ReDim used(1 To lastPeriod)
    Dim list() As Long: ReDim list(1 To lastPeriod)
    Dim count As Long: count = 0
    Dim i As Long, effectivePref As String

    If diff < 0 Then
        effectivePref = "LAST"
    Else
        effectivePref = UCase$(Trim$(pref))
        If effectivePref = "" Then effectivePref = "LAST"
    End If

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
                    n = CLng(val(s))
                    If n >= 1 And n <= lastPeriod Then
                        If Not used(n) Then count = count + 1: list(count) = n: used(n) = True
                    End If
                End If
            Next s
            For i = lastPeriod To 1 Step -1
                If Not used(i) Then count = count + 1: list(count) = i: used(i) = True
            Next i
    End Select

    ReDim Preserve list(1 To count)
    BuildBalanceOrder = list
End Function

Private Function FindScheduleRow() As Long
    Dim i As Long, lastRow As Long
    lastRow = m_wsTienDo.Cells(m_wsTienDo.Rows.count, "C").End(xlUp).Row
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
    ' M?ng: [1]=S? ti?n, [2]=Ngày TT, [3]=% trên T?ng, [4]=S? ngày c?ng
    ReDim m_TienDoThanhToan(1 To lastPeriod, 1 To 4)

    Dim totalPaid As Currency, currentDate As Date
    totalPaid = 0
    currentDate = Me.NgayTTDot1

    Dim pctDot1 As Variant, pctDot2 As Variant
    pctDot1 = m_wsTienDo.Cells(scheduleRow, 5).value   ' E – Ð?t 1
    pctDot2 = m_wsTienDo.Cells(scheduleRow, 7).value   ' G – Ð?t 2

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
            ' ==== Các case d?c bi?t HÐMB – TRÒN 1 L?N ? CU?I ====
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
            ' ==== KHÔNG HÐMB ====
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

    ' ====== CÂN B?NG PH?N DU ======
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

    ' C?p nh?t l?i % tooltip
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


