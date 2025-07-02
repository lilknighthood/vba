Attribute VB_Name = "TT_Canho"
Option Explicit
Sub TinhToanTongHop_ChoDongHienTai()
    '--- KHAI BAO ---
    Dim wsSetup As Worksheet, wsData As Worksheet
    Dim activeRow As Long
    Dim colGiaBan As String, colDtThongThuy As String, colTenTienDo As String, colBatDauNgayTT As String
    Dim colGiaTriCanHo As String, colGiaTriQSDD As String, colThueGTGT As String, colPhiBaoTri As String
    Dim colBC_GiaBan As String, colBC_GiaTriCH As String, colBC_GiaTriQSDD As String, colBC_ThueGTGT As String, colBC_PhiBaoTri As String

    '--- DOC CAU HINH ---
    On Error Resume Next
    Set wsSetup = ThisWorkbook.Sheets("Setup")
    If wsSetup Is Nothing Then
        MsgBox "Khong tim thay sheet 'Setup'.", vbCritical, "Loi cau hinh"
        Exit Sub
    End If
    On Error GoTo 0
    
    With wsSetup
        colGiaBan = .Range("B1").Value: colDtThongThuy = .Range("B2").Value
        colGiaTriCanHo = .Range("B3").Value: colGiaTriQSDD = .Range("B4").Value
        colThueGTGT = .Range("B5").Value: colPhiBaoTri = .Range("B6").Value
        colTenTienDo = .Range("B7").Value
        colBatDauNgayTT = .Range("B9").Value
        colBC_GiaBan = .Range("B10").Value: colBC_GiaTriCH = .Range("B11").Value
        colBC_GiaTriQSDD = .Range("B12").Value: colBC_ThueGTGT = .Range("B13").Value
        colBC_PhiBaoTri = .Range("B14").Value
    End With

    '--- BAT DAU XU LY ---
    activeRow = ActiveCell.Row
    Set wsData = ThisWorkbook.Sheets("CAN HO K-HOME")

    '========================================================================
    '   *** BUOC KIEM TRA DIEU KIEN TRUOC KHI CHAY ***
    '========================================================================
    If wsData.Range(colTenTienDo & activeRow).Value = "" Then
        MsgBox "Ban chua chon TIEN DO THANH TOAN tai dong " & activeRow, vbExclamation, "Thieu thong tin"
        Exit Sub
    End If
    
    If Not IsDate(wsData.Range(colBatDauNgayTT & activeRow).Value) Then
        MsgBox "Ban chua nhap NGAY THANH TOAN DOT 1 tai dong " & activeRow, vbExclamation, "Thieu thong tin"
        Exit Sub
    End If
    
    '========================================================================
    '   PHAN 1: TINH TOAN CAC GIA TRI CO BAN CUA CAN HO
    '========================================================================
    Dim heSoDat As Double
    Dim giaBanCanHo As Currency, dtThongThuy As Double
    Dim giaTriQSDD As Currency, giaTriCanHo As Currency, thueGTGT As Currency, phiBaoTri As Currency
    
    heSoDat = 729754.9204
    giaBanCanHo = wsData.Range(colGiaBan & activeRow).Value
    dtThongThuy = wsData.Range(colDtThongThuy & activeRow).Value
    
    If giaBanCanHo <= 0 Or dtThongThuy <= 0 Then
        MsgBox "Du lieu dau vao (Gia ban hoac DTSD) khong hop le.", vbExclamation
        Exit Sub
    End If
    
    giaTriQSDD = dtThongThuy * heSoDat
    giaTriCanHo = (giaBanCanHo - giaTriQSDD) / 1.1
    thueGTGT = giaTriCanHo * 0.1
    phiBaoTri = (giaTriQSDD + giaTriCanHo) * 0.02
    
    With wsData
        .Range(colGiaTriCanHo & activeRow).Value = giaTriCanHo
        .Range(colGiaTriQSDD & activeRow).Value = giaTriQSDD
        .Range(colThueGTGT & activeRow).Value = thueGTGT
        .Range(colPhiBaoTri & activeRow).Value = phiBaoTri
        
        .Range(colBC_GiaBan & activeRow).Value = vnd(giaBanCanHo)
        .Range(colBC_GiaTriCH & activeRow).Value = vnd(giaTriCanHo)
        .Range(colBC_GiaTriQSDD & activeRow).Value = vnd(giaTriQSDD)
        .Range(colBC_ThueGTGT & activeRow).Value = vnd(thueGTGT)
        .Range(colBC_PhiBaoTri & activeRow).Value = vnd(phiBaoTri)
    End With
    
    '========================================================================
    '   GOI SUB PHU DE THUC HIEN PHAN 2
    '========================================================================
    Call TinhTienDoThanhToan(activeRow, giaBanCanHo)
    
    MsgBox "Hoan tat! Da tinh toan xong cho dong " & activeRow, vbInformation, "Thanh cong"
End Sub
