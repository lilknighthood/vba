Option Explicit
Private Const MAX_PAYMENT_PERIODS As Integer = 15
Private Const SHEET_SETUP As String = "Setup"
Private Const SHEET_TIENDO As String = "TIEN_DO_TT"

'--- DAU VAO ---
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

'--- DAU RA ---
Public SoHopDong As String
Public GiaTriGocDeTinhTienDo As Currency
Public GiaTriDeGhiVaoCotCoc As Currency
Public TiLeThanhToanDot1 As Double

'--- RIENG TU ---
Private m_TienDoThanhToan() As Variant
Private m_wsSetup As Worksheet
Private m_wsTienDo As Worksheet
Private m_totalPct As Double   ' Tong % tien do (dang thap phan, vi du 0.3 = 30%)


'========================================================================
'                       KHOI TAO CLASS
'========================================================================
Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    Set m_wsSetup = ThisWorkbook.Sheets(SHEET_SETUP)
    Set m_wsTienDo = ThisWorkbook.Sheets(SHEET_TIENDO)
    Exit Sub
ErrorHandler:
    MsgBox "LOI KHOI TAO CLASS!" & vbCrLf & Err.Description, vbCritical
    End
End Sub

'========================================================================
'                       THUOC TINH
'========================================================================
Public Property Get TienDoThanhToan() As Variant
    TienDoThanhToan = m_TienDoThanhToan
End Property

Public Property Get IsHDMBContract() As Boolean
    Dim nm As String: nm = UCase$(Me.TenTienDo)
    IsHDMBContract = (InStr(1, nm, "HDMB") > 0) Or (InStr(1, nm, "H" & ChrW(272) & "MB") > 0)
End Property


Public Property Get totalPct() As Double
    totalPct = m_totalPct
End Property

Public Property Let totalPct(ByVal value As Double)
    m_totalPct = value
End Property


'========================================================================
'                       PHUONG THUC CHINH
'========================================================================
' Luon dung THANH_TIEN; luu tong % de xu ly dat cuoi va tinh tien coc
Public Sub XacDinhGiaTriGoc(ByVal totalPercent As Double)
    ' THÀNH_TI?N = THÀNH_TI?N_Ð?T + THÀNH_TI?N_NHÀ
    If Me.ThanhTienDat_Input > 0 Or Me.ThanhTienNha_Input > 0 Then
        Me.TongThanhTien = Me.ThanhTienDat_Input + Me.ThanhTienNha_Input
    End If

    ' Gia tri goc de tinh tien do = TongThanhTien (khong chia ty le nua)
    Me.GiaTriGocDeTinhTienDo = Me.TongThanhTien

    ' Luu lai tong %
    Me.totalPct = totalPercent
End Sub


Public Sub TaoSoHopDong()
    If Len(Trim(Me.MaSoLo)) = 0 Or Not IsDate(Me.NgayKy) Or Len(Trim(Me.TenTienDo)) = 0 Then Exit Sub
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

' THEM HAM MOI - Thiet lap gia tri Dot 1 tu sheet data
Public Sub SetManualDot1Value(ByVal manualDot1 As Currency)
    ' Ham nay se duoc goi tu MainModule de truyen gia tri Dot 1 thu cong
    m_ManualDot1 = manualDot1
End Sub
Public Sub SetManualDot2Value(ByVal manualDot2 As Currency)
    m_ManualDot2 = manualDot2
End Sub


'========================================================================
'                       HAM HO TRO (PRIVATE)
'========================================================================

Private Function GetContractTemplate(ByVal lookupValue As String) As String
    Dim rng As Range, defTpl As String, tpl As String, i As Long
    Dim key As String
    
    On Error Resume Next
    Set rng = m_wsSetup.Range("G2:H" & m_wsSetup.Cells(m_wsSetup.Rows.Count, "G").End(xlUp).row)
    On Error GoTo 0
    
    If rng Is Nothing Or rng.Rows.Count < 1 Then Exit Function
    
    ' Lay mau mac dinh
    defTpl = rng.Cells(rng.Rows.Count, 2).value
    tpl = defTpl
    
    ' Chuyen gia tri sang chu Hoa
    lookupValue = UCase$(lookupValue)
    
    ' Vong lap tim mau Hop Dong
    For i = 1 To rng.Rows.Count - 1
        key = Trim$(UCase$(rng.Cells(i, 1).value))
        
        ' Neu key khong rong va Ten Tien Do Co CHUA Key
        ' thi lay mau tuong ung va thoat
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
    ' Mo rong mang len 4 cot: So tien, Ngay TT, % tinh duoc, So ngay tang
    ReDim m_TienDoThanhToan(1 To lastPeriod, 1 To 4)

    Dim totalPaid As Currency, currentDate As Date
    totalPaid = 0
    currentDate = Me.NgayTTDot1

    Dim pctDot1 As Variant, pctDot2 As Variant
    pctDot1 = m_wsTienDo.Cells(scheduleRow, 5).value     ' Cot E – Dot 1
    pctDot2 = m_wsTienDo.Cells(scheduleRow, 7).value     ' Cot G – Dot 2

    Dim isDot1Empty As Boolean, isDot2Empty As Boolean, isDot1_25pct As Boolean
    isDot1Empty = Not (IsNumeric(pctDot1) And Len(CStr(pctDot1)) > 0)
    isDot2Empty = Not (IsNumeric(pctDot2) And Len(CStr(pctDot2)) > 0)
    isDot1_25pct = (IsNumeric(pctDot1) And Abs(CDbl(pctDot1) - 0.25) < 0.0001)

    Dim i As Integer
    For i = 1 To lastPeriod
        Dim pay As Currency, pct_raw As Variant, days_raw As Variant
        pct_raw = m_wsTienDo.Cells(scheduleRow, (i - 1) * 2 + 5).value
        days_raw = m_wsTienDo.Cells(scheduleRow, (i - 2) * 2 + 6).value

        ' === 3 TRUONG HOP HÐMB ===
        If Me.IsHDMBContract Then
            If i = 1 And isDot1Empty Then
                pay = m_ManualDot1

            ElseIf i = 2 And isDot1Empty Then
                pay = Round(Me.TongThanhTien * 0.7, 0) - m_ManualDot1
                If pay < 0 Then pay = 0

            ElseIf i = 2 And isDot1_25pct And isDot2Empty Then
                pay = m_ManualDot2

            ElseIf i = 3 And isDot1_25pct And isDot2Empty Then
                pay = Round(Me.TongThanhTien * 0.7, 0) - _
                      Round(Me.TongThanhTien * 0.25, 0) - m_ManualDot2
                If pay < 0 Then pay = 0

            Else
                If IsNumeric(pct_raw) Then
                    pay = Round(Me.TongThanhTien * CDbl(pct_raw), 0)
                Else
                    pay = 0
                End If
            End If

        Else
            ' Truong hop binh thuong
            If IsNumeric(pct_raw) Then
                pay = Round(Me.TongThanhTien * CDbl(pct_raw), 0)
            Else
                pay = 0
            End If
        End If

        totalPaid = totalPaid + pay
        m_TienDoThanhToan(i, 1) = pay ' So tien

        ' Ngay thanh toan
        If i = 1 Then
            m_TienDoThanhToan(i, 2) = currentDate
        ElseIf IsNumeric(days_raw) And Len(CStr(days_raw)) > 0 Then
            currentDate = DateAdd("d", CLng(days_raw), currentDate)
            m_TienDoThanhToan(i, 2) = currentDate
        Else
            m_TienDoThanhToan(i, 2) = ""
        End If

        ' % tinh duoc
        If Me.TongThanhTien > 0 Then
            m_TienDoThanhToan(i, 3) = Round(pay / Me.TongThanhTien, 6)
        Else
            m_TienDoThanhToan(i, 3) = 0
        End If

        ' Ngay cong (de tooltip)
        m_TienDoThanhToan(i, 4) = days_raw
    Next i
End Sub

