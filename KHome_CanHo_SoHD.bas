Public Sub TaoSoHopDong(ByVal activeRow As Long)
    '--- KHAI BAO ---
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
    
    '========================================================================
    '   *** BUOC 1: TIM DUNG "CONG THUC MAU" TU BANG TRA CUU ***
    '========================================================================
    Dim bangTraCuu As Range, dieuKien As Range
    Dim mauHopDong As String, mauMacDinh As String
    
    'Xac dinh vung chua Bang Tra Cuu (tu G2 den cot H dong cuoi cung co du lieu)
    Set bangTraCuu = wsSetup.Range("G2:H" & wsSetup.Cells(wsSetup.Rows.Count, "G").End(xlUp).row)
    
    'Lay mau hop dong mac dinh (o dong cuoi cung cua bang)
    mauMacDinh = bangTraCuu.Cells(bangTraCuu.Rows.Count, 2).Value
    mauHopDong = mauMacDinh 'Gan gia tri mac dinh truoc
    
    'Duyet qua tung dieu kien trong Bang Tra Cuu (tru dong mac dinh)
    For Each dieuKien In bangTraCuu.Resize(bangTraCuu.Rows.Count - 1).Rows
        Dim tuKhoa As String
        tuKhoa = dieuKien.Cells(1, 1).Value 'Lay tu khoa o cot G
        
        'Kiem tra xem tu khoa co phai la chuoi "HÐMB" khong
        If UCase(tuKhoa) = "HÐMB" Or UCase(tuKhoa) = "H" & ChrW(272) & "MB" Then
            tuKhoa = "*" & "H" & ChrW(272) & "MB" & "*" 'Chuan hoa de so sanh voi Like
        Else
            tuKhoa = "*" & tuKhoa & "*" 'Them dau * de tim kiem o bat ky vi tri nao
        End If

        If tenTienDo Like tuKhoa Then
            mauHopDong = dieuKien.Cells(1, 2).Value 'Lay mau hop dong tuong ung o cot H
            Exit For 'Tim thay roi, thoat khoi vong lap
        End If
    Next dieuKien

    '========================================================================
    '   *** BUOC 2: TAO SO HOP DONG TU "CONG THUC MAU" DA TIM DUOC ***
    '========================================================================
    Dim soHopDongHoanChinh As String
    soHopDongHoanChinh = mauHopDong
    
    'Thay the [NAMKY]
    soHopDongHoanChinh = Replace(soHopDongHoanChinh, "[NAMKY]", Year(ngayKy), 1, -1, vbTextCompare)
    
    'Thay the [CANHO]
    soHopDongHoanChinh = Replace(soHopDongHoanChinh, "[CANHO]", maCanHo, 1, -1, vbTextCompare)
    
    'Ghi ket qua cuoi cung vao cot So Hop Dong
    wsData.Range(colSoHD & activeRow).Value = soHopDongHoanChinh
End Sub
