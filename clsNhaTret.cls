Option Explicit

'======================================================='
'   CLASS MODULE: clsNhaTret - Phien ban sua loi cu phap  '
'======================================================='

' --- HANG SO ---
Private Const MAX_PAYMENT_PERIODS As Integer = 20
Private Const HDMB_KEYWORD As String = "HÐMB"
Private Const SHEET_SETUP As String = "Setup"
Private Const SHEET_TIENDO As String = "TIEN_DO_TT"

' --- CAC THUOC TINH PUBLIC (DAU VAO) ---
Public RowNum As Long
Public TongThanhTien As Currency
Public MaSoLo As String
Public NgayKy As Date
Public TenTienDo As String
Public NgayTTDot1 As Date
Public ThanhTienDat_Input As Currency
Public ThanhTienNha_Input As Currency

' --- CAC THUOC TINH PUBLIC (DAU RA) ---
Public SoHopDong As String
Public GiaTriGocDeTinhTienDo As Currency
Public GiaTriDeGhiVaoCotCoc As Currency
Public TiLeThanhToanDot1 As Double

' --- BIEN THANH VIEN PRIVATE ---
Private m_TienDoThanhToan() As Variant
Private m_wsSetup As Worksheet
Private m_wsTienDo As Worksheet

'========================================================================
'                       KHOI TAO CLASS
'========================================================================

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    Set m_wsSetup = ThisWorkbook.Sheets(SHEET_SETUP)
    Set m_wsTienDo = ThisWorkbook.Sheets(SHEET_TIENDO)
    Exit Sub

ErrorHandler:
    MsgBox "LOI KHOI TAO CLASS!" & vbCrLf & vbCrLf & _
           "Khong the tim thay sheet co ten: """ & Err.Description & """." & vbCrLf & _
           "Vui long kiem tra lai ten sheet trong file Excel.", vbCritical, "Loi Thiet Lap Sheet"
    End
End Sub

'========================================================================
'                       CAC THUOC TINH CHI DOC
'========================================================================

Public Property Get TienDoThanhToan() As Variant
    TienDoThanhToan = m_TienDoThanhToan
End Property

Public Property Get IsHDMBContract() As Boolean
    Dim upperTenTienDo As String
    upperTenTienDo = UCase(Me.TenTienDo)
    IsHDMBContract = (InStr(1, upperTenTienDo, "HÐMB") > 0) Or (InStr(1, upperTenTienDo, "HDMB") > 0)
End Property

'========================================================================
'                       CAC PHUONG THUC PUBLIC CHINH
'========================================================================

Public Sub XacDinhGiaTriGoc(ByVal totalPercentage As Double)
    If Me.IsHDMBContract Then
        Me.GiaTriGocDeTinhTienDo = Me.TongThanhTien
        Me.GiaTriDeGhiVaoCotCoc = 0
    Else
        Me.GiaTriGocDeTinhTienDo = VBA.Round(Me.TongThanhTien * totalPercentage, 0)
        Me.GiaTriDeGhiVaoCotCoc = Me.GiaTriGocDeTinhTienDo
    End If
End Sub

Public Sub TaoSoHopDong()
    If Len(Trim(Me.MaSoLo)) = 0 Or Not IsDate(Me.NgayKy) Or Len(Trim(Me.TenTienDo)) = 0 Then Exit Sub
    Dim contractTemplate As String: contractTemplate = GetContractTemplate(Me.TenTienDo)
    If Len(contractTemplate) = 0 Then Exit Sub
    Dim result As String: result = contractTemplate
    result = Replace(result, "[NAMKY]", CStr(Year(Me.NgayKy)), 1, -1, vbTextCompare)
    result = Replace(result, "[LO-O]", Me.MaSoLo, 1, -1, vbTextCompare)
    Me.SoHopDong = result
End Sub

Public Sub TinhTienDoThanhToan()
    If Me.GiaTriGocDeTinhTienDo <= 0 Then Exit Sub
    Dim scheduleRow As Long: scheduleRow = FindScheduleRow()
    If scheduleRow = 0 Then Exit Sub
    
    '======================================================='
    '            *** SUA LOI CU PHAP TAI DAY ***
    ' Chuyen cau lenh If...Then...Else thanh mot khoi da dong
    '======================================================='
    Dim tiLeDot1Value As Variant
    tiLeDot1Value = m_wsTienDo.Cells(scheduleRow, 5).Value
    If IsNumeric(tiLeDot1Value) Then
        Me.TiLeThanhToanDot1 = CDbl(tiLeDot1Value)
    Else
        Me.TiLeThanhToanDot1 = 0
    End If
    '======================================================='
    '                KET THUC THAY DOI
    '======================================================='

    Dim lastPeriod As Integer: lastPeriod = FindLastPaymentPeriod(scheduleRow)
    If lastPeriod = 0 Then Exit Sub
    
    BuildPaymentSchedule scheduleRow, lastPeriod
End Sub

'========================================================================
'                       CAC PHUONG THUC PHU TRO (PRIVATE)
'========================================================================

Private Function GetContractTemplate(ByVal lookupValue As String) As String
    Dim lookupRange As Range, defaultTemplate As String, contractTemplate As String, i As Long
    On Error Resume Next
    Set lookupRange = m_wsSetup.Range("G2:H" & m_wsSetup.Cells(m_wsSetup.Rows.Count, "G").End(xlUp).row)
    On Error GoTo 0
    If lookupRange Is Nothing Or lookupRange.Rows.Count < 1 Then Exit Function
    defaultTemplate = lookupRange.Cells(lookupRange.Rows.Count, 2).Value
    contractTemplate = defaultTemplate
    For i = 1 To lookupRange.Rows.Count - 1
        If Trim(UCase(lookupRange.Cells(i, 1).Value)) = Trim(UCase(lookupValue)) Then
            contractTemplate = lookupRange.Cells(i, 2).Value
            Exit For
        End If
    Next i
    GetContractTemplate = contractTemplate
End Function

Private Function FindScheduleRow() As Long
    Dim lastRow As Long, i As Long
    lastRow = m_wsTienDo.Cells(m_wsTienDo.Rows.Count, "C").End(xlUp).row
    For i = 1 To lastRow
        If Trim(UCase(m_wsTienDo.Cells(i, "C").Value)) = Trim(UCase(Me.TenTienDo)) Then
            FindScheduleRow = i
            Exit Function
        End If
    Next i
End Function

Private Function FindLastPaymentPeriod(ByVal scheduleRow As Long) As Integer
    Dim i As Integer, colIndex As Integer, cellValue As Variant
    For i = MAX_PAYMENT_PERIODS To 1 Step -1
        colIndex = (i - 1) * 2 + 5
        cellValue = m_wsTienDo.Cells(scheduleRow, colIndex).Value
        If IsNumeric(cellValue) And Len(CStr(cellValue)) > 0 Then
            FindLastPaymentPeriod = i
            Exit Function
        End If
    Next i
End Function

Private Sub BuildPaymentSchedule(ByVal scheduleRow As Long, ByVal lastPeriod As Integer)
    ReDim m_TienDoThanhToan(1 To lastPeriod, 1 To 2)
    Dim totalPaid As Currency, currentPaymentDate As Date, i As Integer
    totalPaid = 0: currentPaymentDate = Me.NgayTTDot1
    
    For i = 1 To lastPeriod
        Dim paymentAmount As Currency
        If i < lastPeriod Then
            Dim percentage As Variant
            percentage = m_wsTienDo.Cells(scheduleRow, (i - 1) * 2 + 5).Value
            If IsNumeric(percentage) Then
                paymentAmount = VBA.Round(Me.GiaTriGocDeTinhTienDo * CDbl(percentage), 0)
            Else
                paymentAmount = 0
            End If
            totalPaid = totalPaid + paymentAmount
        Else
            paymentAmount = Me.GiaTriGocDeTinhTienDo - totalPaid
        End If
        m_TienDoThanhToan(i, 1) = paymentAmount
        
        If i = 1 Then
            m_TienDoThanhToan(i, 2) = currentPaymentDate
        Else
            Dim daysToAdd As Variant
            daysToAdd = m_wsTienDo.Cells(scheduleRow, (i - 2) * 2 + 6).Value
            If IsNumeric(daysToAdd) And Len(CStr(daysToAdd)) > 0 Then
                currentPaymentDate = DateAdd("d", CLng(daysToAdd), currentPaymentDate)
                m_TienDoThanhToan(i, 2) = currentPaymentDate
            Else
                m_TienDoThanhToan(i, 2) = ""
            End If
        End If
    Next i
End Sub


