Option Explicit
Sub TinhTienDoThanhToan(ByVal activeRow As Long, ByVal giaBanCanHo As Currency, ByVal giaTriCanHo As Currency)
    '--- KHAI BAO ---
    Dim wsSetup As Worksheet, wsData As Worksheet, wsTienDo As Worksheet
    Dim colTenTienDo As String, colBatDauTraTien As String, colBatDauNgayTT As String, colBC_Dot1 As String
    Dim colTiLeTT As String, colTienCoc As String, colKiemTra As String

    '--- DOC CAU HINH ---
    Set wsSetup = ThisWorkbook.Sheets("Setup"): Set wsData = ThisWorkbook.Sheets("CAN HO K-HOME"): Set wsTienDo = ThisWorkbook.Sheets("TIEN_DO_TT")
    With wsSetup
        colTenTienDo = .Range("B7").Value: colBatDauTraTien = .Range("B8").Value
        colBatDauNgayTT = .Range("B9").Value: colBC_Dot1 = .Range("B15").Value
        colTiLeTT = .Range("B16").Value: colTienCoc = .Range("B20").Value
        colKiemTra = .Range("B21").Value
    End With

    Dim tenTienDoCanTim As String: tenTienDoCanTim = wsData.Range(colTenTienDo & activeRow).Value
    If tenTienDoCanTim = "" Then Exit Sub
    
    '--- Tim dong tien do ---
    Dim dongTienDo As Long, timThay As Boolean, i As Integer: timThay = False
    For dongTienDo = 1 To wsTienDo.Cells(wsTienDo.Rows.Count, "C").End(xlUp).row
        If wsTienDo.Cells(dongTienDo, "C").Value = tenTienDoCanTim Then timThay = True: Exit For
    Next dongTienDo
    If Not timThay Then Exit Sub
    
    '--- Tinh Tong Ty Le % ---
    Dim tongTyLePhanTram As Double: tongTyLePhanTram = 0
    For i = 1 To 16
        If IsNumeric(wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value) Then
            tongTyLePhanTram = tongTyLePhanTram + wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value
        End If
    Next i
    
    '--- Tinh va ghi gia tri vao cot %_TIEN_COC ---
    Dim tienCocValue As Currency
    tienCocValue = giaTriCanHo * tongTyLePhanTram
    wsData.Range(colTienCoc & activeRow).Value = tienCocValue
    
    '========================================================================
    '   *** LOGIC CHINH: XAC DINH GIA TRI NEN DE TINH TOAN ***
    '========================================================================
    Dim baseAmount As Currency
    If tenTienDoCanTim Like "*" & "H" & ChrW(272) & "MB" & "*" Then
        baseAmount = giaBanCanHo 'Neu la HÐMB, tinh tren Gia Ban
    Else
        baseAmount = tienCocValue 'Neu la Coc, tinh tren gia tri Tien Coc
    End If
    
    '--- Tim dot cuoi cung ---
    Dim dotCuoiCung As Integer: dotCuoiCung = 0
    For i = 16 To 1 Step -1
        If IsNumeric(wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value) And wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value <> "" Then
            dotCuoiCung = i: Exit For
        End If
    Next i
    
    '--- Don dep du lieu truoc khi tinh ---
    Dim colIndexOutput As Long, colIndexBC_Dot1 As Long, colIndexNgayOutput As Long
    colIndexOutput = wsData.Range(colBatDauTraTien & 1).Column
    colIndexBC_Dot1 = wsData.Range(colBC_Dot1 & 1).Column
    colIndexNgayOutput = wsData.Range(colBatDauNgayTT & 1).Column
    For i = 1 To 16
        wsData.Cells(activeRow, colIndexOutput + (i - 1) * 2).ClearContents
        wsData.Cells(activeRow, colIndexBC_Dot1 + i - 1).ClearContents
        If i > 1 Then wsData.Cells(activeRow, colIndexNgayOutput + (i - 1) * 2).ClearContents
    Next i
    
    If dotCuoiCung = 0 Then wsData.Range(colKiemTra & activeRow).ClearContents: Exit Sub 'Thoat neu khong co dot nao

    '--- Tinh toan cac dot ---
    Dim tongTienDaTra As Currency, soTienPhaiTra As Currency, tyLePhanTram As Variant
    Dim ngayThanhToanHienTai As Date, ngayTiepTheo As Date, soNgayCongThem As Variant
    
    tongTienDaTra = 0
    ngayThanhToanHienTai = wsData.Cells(activeRow, colIndexNgayOutput).Value
    
    For i = 1 To dotCuoiCung
        If i < dotCuoiCung Then
            tyLePhanTram = wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value
            'Luu y: ty le % luon tinh tren Gia Ban Can Ho
            soTienPhaiTra = VBA.Round(giaBanCanHo * tyLePhanTram, 0)
            tongTienDaTra = tongTienDaTra + soTienPhaiTra
        Else
            soTienPhaiTra = baseAmount - tongTienDaTra
        End If
        
        wsData.Cells(activeRow, colIndexOutput + (i - 1) * 2).Value = soTienPhaiTra
        wsData.Cells(activeRow, colIndexBC_Dot1 + i - 1).Value = vnd(soTienPhaiTra)
        
        If i > 1 Then
            soNgayCongThem = wsTienDo.Cells(dongTienDo, (i - 2) * 2 + 6).Value
            If IsNumeric(soNgayCongThem) And soNgayCongThem <> "" Then
                ngayTiepTheo = DateAdd("d", soNgayCongThem, ngayThanhToanHienTai)
                wsData.Cells(activeRow, colIndexNgayOutput + (i - 1) * 2).Value = ngayTiepTheo
                ngayThanhToanHienTai = ngayTiepTheo
            End If
        End If
    Next i
    
    'Ghi gia tri cho cot Kiem Tra bang tong so tien thuc te da tra
    wsData.Range(colKiemTra & activeRow).Value = baseAmount
End Sub
