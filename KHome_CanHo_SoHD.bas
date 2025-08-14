
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

    '--- KHOI TAO ---
    Set wsData = ThisWorkbook.Sheets("CAN HO K-HOME")

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
    soHopDongHoanChinh = Replace(soHopDongHoanChinh, "[CANHO]", maCanHo, 1, -1, vbTextCompare)
    
    wsData.Range(colSoHD & activeRow).Value = soHopDongHoanChinh
End Sub
