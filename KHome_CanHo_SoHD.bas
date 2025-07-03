Option Explicit
Sub TaoSoHopDong(ByVal activeRow As Long)
    '--- KHAI BAO ---
    Dim wsSetup As Worksheet, wsData As Worksheet
    
    '--- KHAI BAO CAU HINH ---
    Dim colCanHo As String
    Dim colNgayKy As String
    Dim colSoHD As String
    Dim colTenTienDo As String 'De lay ten tien do

    '--- DOC CAU HINH ---
    On Error Resume Next
    Set wsSetup = ThisWorkbook.Sheets("Setup")
    If wsSetup Is Nothing Then Exit Sub 'Thoat neu khong co sheet Setup
    
    With wsSetup
        colCanHo = .Range("B17").Value
        colNgayKy = .Range("B18").Value
        colSoHD = .Range("B19").Value
        colTenTienDo = .Range("B7").Value 'Doc them cau hinh cot Tien Do
    End With
    
    If colCanHo = "" Or colNgayKy = "" Or colSoHD = "" Or colTenTienDo = "" Then Exit Sub 'Thoat neu thieu cau hinh

    '--- KHOI TAO ---
    Set wsData = ThisWorkbook.Sheets("CAN HO K-HOME")

    '--- Lay du lieu dau vao ---
    Dim maCanHo As String
    Dim ngayKy As Variant
    Dim tenTienDo As String
    Dim chuoiCanTim As String '*** BIEN MOI: De dung ChrW ***
    
    maCanHo = wsData.Range(colCanHo & activeRow).Value
    ngayKy = wsData.Range(colNgayKy & activeRow).Value
    tenTienDo = wsData.Range(colTenTienDo & activeRow).Value
    
    '*** TAO CHUOI CAN TIM VOI KY TU "Ð" CHUAN UNICODE ***
    chuoiCanTim = "H" & ChrW(272) & "MB" 'Tao ra chu "HÐMB"
    
    '--- Kiem tra dieu kien truoc khi tao so hop dong ---
    If maCanHo <> "" And IsDate(ngayKy) Then
        
        '*** KIEM TRA BANG CHUOI UNICODE DA TAO ***
        If tenTienDo Like "*" & chuoiCanTim & "*" Then
            'Neu ten tien do co chua chuoi "HÐMB", dung dinh dang VAY
            wsData.Range(colSoHD & activeRow).Value = maCanHo & "/" & Year(ngayKy) & "/2025-HÐMBVAY"
        Else
            'Neu khong, dung dinh dang thong thuong
            wsData.Range(colSoHD & activeRow).Value = maCanHo & "/" & Year(ngayKy) & "/2025-HÐMB"
        End If
        
    End If
End Sub
