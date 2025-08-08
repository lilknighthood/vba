'======================================================='
'   CLASS MODULE: clsNhaTret (logic m?i)               '
'======================================================='

Private Const MAX_PAYMENT_PERIODS As Integer = 20
Private Const SHEET_SETUP As String = "Setup"
Private Const SHEET_TIENDO As String = "TIEN_DO_TT"

'--- INPUT ---
Public RowNum As Long
Public TongThanhTien As Currency
Public MaSoLo As String
Public NgayKy As Date
Public TenTienDo As String
Public NgayTTDot1 As Date
Public ThanhTienDat_Input As Currency
Public ThanhTienNha_Input As Currency

'--- OUTPUT ---
Public SoHopDong As String
Public GiaTriGocDeTinhTienDo As Currency
Public GiaTriDeGhiVaoCotCoc As Currency
Public TiLeThanhToanDot1 As Double

'--- PRIVATE ---
Private m_TienDoThanhToan() As Variant
Private m_wsSetup As Worksheet
Private m_wsTienDo As Worksheet
Private m_TotalPct As Double   ' T?ng % ti?n d? (d?ng th?p phân)

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    Set m_wsSetup = ThisWorkbook.Sheets(SHEET_SETUP)
    Set m_wsTienDo = ThisWorkbook.Sheets(SHEET_TIENDO)
    Exit Sub
ErrorHandler:
    MsgBox "LOI KHOI TAO CLASS!" & vbCrLf & Err.Description, vbCritical
    End
End Sub

Public Property Get TienDoThanhToan() As Variant
    TienDoThanhToan = m_TienDoThanhToan
End Property

Public Property Get IsHDMBContract() As Boolean
    Dim nm As String
    nm = UCase$(Me.TenTienDo)
    IsHDMBContract = (InStr(1, nm, "HÐMB") > 0) Or (InStr(1, nm, "HDMB") > 0)
End Property

'=== LOGIC M?I ===
Public Sub XacDinhGiaTriGoc(ByVal totalPercentage As Double)
    ' Luôn dùng THÀNH_TI?N
    Me.GiaTriGocDeTinhTienDo = Me.TongThanhTien
    
    ' Luu t?ng % cho công th?c d?t cu?i
    m_TotalPct = totalPercentage
    
    ' Giá tr? ghi c?t C?C: HÐMB = 0, ngu?c l?i = THÀNH_TI?N
    If Me.IsHDMBContract Then
        Me.GiaTriDeGhiVaoCotCoc = 0
    Else
        Me.GiaTriDeGhiVaoCotCoc = Me.TongThanhTien
    End If
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
    
    Dim v As Variant: v = m_wsTienDo.Cells(rowSch, 5).Value
    If IsNumeric(v) Then Me.TiLeThanhToanDot1 = CDbl(v) Else Me.TiLeThanhToanDot1 = 0
    
    Dim lastPeriod As Integer: lastPeriod = FindLastPaymentPeriod(rowSch)
    If lastPeriod = 0 Then Exit Sub
    
    BuildPaymentSchedule rowSch, lastPeriod
End Sub

'----------------- PRIVATE HELPERS ----------------------

Private Function GetContractTemplate(ByVal lookupValue As String) As String
    Dim rng As Range, defTpl As String, tpl As String, i As Long
    On Error Resume Next
    Set rng = m_wsSetup.Range("G2:H" & m_wsSetup.Cells(m_wsSetup.Rows.Count, "G").End(xlUp).row)
    On Error GoTo 0
    If rng Is Nothing Or rng.Rows.Count < 1 Then Exit Function
    defTpl = rng.Cells(rng.Rows.Count, 2).Value
    tpl = defTpl
    For i = 1 To rng.Rows.Count - 1
        If Trim$(UCase$(rng.Cells(i, 1).Value)) = Trim$(UCase$(lookupValue)) Then
            tpl = rng.Cells(i, 2).Value
            Exit For
        End If
    Next i
    GetContractTemplate = tpl
End Function

Private Function FindScheduleRow() As Long
    Dim i As Long, lastRow As Long
    lastRow = m_wsTienDo.Cells(m_wsTienDo.Rows.Count, "C").End(xlUp).row
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
    ReDim m_TienDoThanhToan(1 To lastPeriod, 1 To 2)
    Dim totalPaid As Currency, currentDate As Date, i As Integer
    totalPaid = 0
    currentDate = Me.NgayTTDot1
    
    For i = 1 To lastPeriod
        Dim pay As Currency
        If i < lastPeriod Then
            Dim pct As Variant
            pct = m_wsTienDo.Cells(scheduleRow, (i - 1) * 2 + 5).Value
            If IsNumeric(pct) Then
                pay = VBA.Round(Me.GiaTriGocDeTinhTienDo * CDbl(pct), 0)
            Else
                pay = 0
            End If
            totalPaid = totalPaid + pay
        Else
            ' --- Ð?T CU?I: (THÀNH_TI?N * T?NG %) - các d?t tru?c ---
            Dim targetTotal As Currency
            targetTotal = VBA.Round(Me.TongThanhTien * m_TotalPct, 0)
            pay = targetTotal - totalPaid
            If pay < 0 Then pay = 0
        End If
        
        m_TienDoThanhToan(i, 1) = pay
        
        ' Ngày
        If i = 1 Then
            m_TienDoThanhToan(i, 2) = currentDate
        Else
            Dim dAdd As Variant
            dAdd = m_wsTienDo.Cells(scheduleRow, (i - 2) * 2 + 6).Value
            If IsNumeric(dAdd) And Len(CStr(dAdd)) > 0 Then
                currentDate = DateAdd("d", CLng(dAdd), currentDate)
                m_TienDoThanhToan(i, 2) = currentDate
            Else
                m_TienDoThanhToan(i, 2) = ""
            End If
        End If
    Next i
End Sub

