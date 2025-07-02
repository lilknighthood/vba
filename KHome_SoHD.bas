Sub TaoSoHopDong(ByVal activeRow As Long)
    '--- KHAI BAO ---
    Dim wsSetup As Worksheet, wsData As Worksheet
    
    '--- KHAI BAO CAU HINH ---
    Dim colCanHo As String
    Dim colNgayKy As String
    Dim colSoHD As String

    '--- DOC CAU HINH ---
    On Error Resume Next
    Set wsSetup = ThisWorkbook.Sheets("Setup")
    If wsSetup Is Nothing Then Exit Sub 'Thoat neu khong co sheet Setup
    
    With wsSetup
        colCanHo = .Range("B17").Value
        colNgayKy = .Range("B18").Value
        colSoHD = .Range("B19").Value
    End With
    
    If colCanHo = "" Or colNgayKy = "" Or colSoHD = "" Then Exit Sub 'Thoat neu thieu cau hinh

    '--- KHOI TAO ---
    Set wsData = ThisWorkbook.Sheets("CAN HO K-HOME")

    '--- Lay du lieu dau vao ---
    Dim maCanHo As String
    Dim ngayKy As Variant
    
    maCanHo = wsData.Range(colCanHo & activeRow).Value
    ngayKy = wsData.Range(colNgayKy & activeRow).Value
    
    '--- Kiem tra dieu kien truoc khi tao so hop dong ---
    If maCanHo <> "" And IsDate(ngayKy) Then
        'Tao chuoi so hop dong theo dung dinh dang
        wsData.Range(colSoHD & activeRow).Value = maCanHo & "/" & Year(ngayKy) & "/2025-HDMB"
    End If
End Sub
=======
Sub TaoSoHopDong()
    '--- KHAI BAO ---
    Dim wsSetup As Worksheet, wsData As Worksheet
    Dim row As Range
    Dim activeRow As Long
    
    '--- KHAI BAO CAU HINH ---
    Dim colCanHo As String
    Dim colNgayKy As String
    Dim colSoHD As String

    '--- DOC CAU HINH ---
    On Error Resume Next
    Set wsSetup = ThisWorkbook.Sheets("Setup")
    If wsSetup Is Nothing Then
        MsgBox "Khong tim thay sheet 'Setup'.", vbCritical, "Loi cau hinh"
        Exit Sub
    End If
    On Error GoTo 0
    
    With wsSetup
        colCanHo = .Range("B17").Value
        colNgayKy = .Range("B18").Value
        colSoHD = .Range("B19").Value
    End With
    
    If colCanHo = "" Or colNgayKy = "" Or colSoHD = "" Then
        MsgBox "Vui long dien day du cau hinh tu B17 den B19 trong sheet 'Setup'.", vbExclamation, "Thieu cau hinh"
        Exit Sub
    End If

    '--- KHOI TAO ---
    Set wsData = ThisWorkbook.Sheets("CAN HO K-HOME")
    Application.ScreenUpdating = False

    '*** VONG LAP QUA CAC DONG DANG HIEN THI TRONG VUNG CHON ***
    For Each row In Selection.Rows
        If row.Hidden = False Then
            activeRow = row.row
            
            '--- Lay du lieu dau vao ---
            Dim maCanHo As String
            Dim ngayKy As Variant
            
            maCanHo = wsData.Range(colCanHo & activeRow).Value
            ngayKy = wsData.Range(colNgayKy & activeRow).Value
            
            '--- Kiem tra dieu kien truoc khi tao so hop dong ---
            If maCanHo <> "" And IsDate(ngayKy) Then
                Dim namKy As Integer
                Dim soHopDong As String
                
                'Lay nam tu ngay ky
                namKy = Year(ngayKy)
                
                'Tao chuoi so hop dong theo dung dinh dang
                soHopDong = maCanHo & "/" & namKy & "/HĐMB"
                
                'Ghi ket qua vao cot So Hop Dong
                wsData.Range(colSoHD & activeRow).Value = soHopDong
            End If
        End If
    Next row
    
    Application.ScreenUpdating = True
    MsgBox "Hoan tat viec tao so hop dong cho cac dong da chon.", vbInformation, "Thanh cong"
End Sub
