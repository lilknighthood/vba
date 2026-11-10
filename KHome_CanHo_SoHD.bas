
'====================================================================================================
' CHUC NANG: Tu dong tao so hop dong
'====================================================================================================
Public Sub TaoSoHopDong(ByVal activeRow As Long)
    '--- KHAI BAO DOI TUONG VA BIEN CAU HINH ---
    Dim wsSetup As Worksheet, wsData As Worksheet
    Dim colCanHo As String, colNgayKy As String, colSoHD As String, colTenTienDo As String
    
    '--- DOC CAU HINH ---
    On Error Resume Next
    Set wsSetup = ThisWorkbook.Sheets("Setup")
    If wsSetup Is Nothing Then Exit Sub
    
    With wsSetup
        colCanHo = .Range("B17").Value
        colNgayKy = .Range("B18").Value
        colSoHD = .Range("B19").Value
        colTenTienDo = .Range("B7").Value
    End With
    
    If colCanHo = "" Or colNgayKy = "" Or colSoHD = "" Or colTenTienDo = "" Then Exit Sub

    
    '======================================================================
    ' T? Ð?NG Ð?C TÊN SHEET DATA T? Ô K2 TRÊN SETUP
    '======================================================================
    Dim tenSheetData As String
    tenSheetData = Trim(wsSetup.Range("K2").Value)
    
    ' N?u b?n quên chua nh?p ? K2, code s? t? l?y tên m?c d?nh cu d? không b? l?i
    If tenSheetData = "" Then tenSheetData = "CAN HO K-HOME"
    
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets(tenSheetData)
    On Error GoTo 0
    
    ' B?y l?i: N?u tên sheet nh?p ? K2 không kh?p v?i b?t k? tab nào bên du?i
    If wsData Is Nothing Then
        MsgBox "Không tìm th?y Sheet d? li?u: [" & tenSheetData & "]." & vbCrLf & _
               "Vui lòng ki?m tra l?i tên sheet ? ô K2 bên sheet SETUP!", vbCritical, "L?i Tên Sheet"
        Exit Sub
    End If
    '======================================================================

    '--- Lay du lieu dau vao ---
    Dim maCanHo As String, ngayKy As Variant, tenTienDo As String
    maCanHo = wsData.Range(colCanHo & activeRow).Value
    ngayKy = wsData.Range(colNgayKy & activeRow).Value
    tenTienDo = wsData.Range(colTenTienDo & activeRow).Value
    
    If maCanHo = "" Or Not IsDate(ngayKy) Then Exit Sub
    
    '--- BUOC 1: TIM DUNG "CONG THUC MAU" TU BANG TRA CUU ---
    Dim bangTraCuu As Range, dieuKien As Range
    Dim mauHopDong As String, mauMacDinh As String
    
    Set bangTraCuu = wsSetup.Range("G2:H" & wsSetup.Cells(wsSetup.Rows.Count, "G").End(xlUp).row)
    mauMacDinh = bangTraCuu.Cells(bangTraCuu.Rows.Count, 2).Value
    mauHopDong = mauMacDinh
    
    For Each dieuKien In bangTraCuu.Resize(bangTraCuu.Rows.Count - 1).Rows
        Dim tuKhoa As String
        tuKhoa = "*" & dieuKien.Cells(1, 1).Value & "*" 'Lay tu khoa va them wildcard

        If UCase(tenTienDo) Like UCase(tuKhoa) Then
            mauHopDong = dieuKien.Cells(1, 2).Value
            Exit For
        End If
    Next dieuKien

    '--- BUOC 2: TAO SO HOP DONG TU "CONG THUC MAU" ---
    Dim soHopDongHoanChinh As String
    soHopDongHoanChinh = mauHopDong
    
    soHopDongHoanChinh = Replace(soHopDongHoanChinh, "[NAMKY]", Year(ngayKy), 1, -1, vbTextCompare)
    soHopDongHoanChinh = Replace(soHopDongHoanChinh, "[NGAYKY]", Format(ngayKy, "dd"), 1, -1, vbTextCompare)
    soHopDongHoanChinh = Replace(soHopDongHoanChinh, "[THANGKY]", Format(ngayKy, "MM"), 1, -1, vbTextCompare)
    soHopDongHoanChinh = Replace(soHopDongHoanChinh, "[NAMKY2]", Format(ngayKy, "yy"), 1, -1, vbTextCompare)
    soHopDongHoanChinh = Replace(soHopDongHoanChinh, "[CANHO]", maCanHo, 1, -1, vbTextCompare)
    
    wsData.Range(colSoHD & activeRow).Value = soHopDongHoanChinh
End Sub

