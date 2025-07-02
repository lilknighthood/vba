Attribute VB_Name = "KHome_NhaTret_TienDoThanhToan"
Option Explicit

Public Sub TinhTienDoThanhToan(ByVal activeRow As Long)
    '--- KHAI BAO ---
    Dim wsSetup As Worksheet, wsData As Worksheet, wsTienDo As Worksheet
    Dim colTenTienDo As String, colBatDauTraTien As String, colBatDauNgayTT As String, colBC_Dot1 As String
    Dim colTiLeTT As String

    '--- DOC CAU HINH (DA CAP NHAT) ---
    Set wsSetup = ThisWorkbook.Sheets("Setup")
    Set wsData = ThisWorkbook.Sheets("FILE TONG HOA PHU - K HOME")
    Set wsTienDo = ThisWorkbook.Sheets("TIEN_DO_TT")
    
    With wsSetup
        colTenTienDo = .Range("B4").Value
        colBatDauTraTien = .Range("B5").Value
        colBatDauNgayTT = .Range("B6").Value
        colBC_Dot1 = .Range("B9").Value
        colTiLeTT = .Range("B10").Value
    End With

    '========================================================================
    '   PHAN 2: TINH TOAN TIEN DO THANH TOAN
    '========================================================================
    Dim tenTienDoCanTim As String
    tenTienDoCanTim = wsData.Range(colTenTienDo & activeRow).Value
    If tenTienDoCanTim = "" Then Exit Sub
    
    '--- Doc gia tri thanh tien tu cot Q, R va tinh tong vao cot S ---
    Dim thanhTienDat As Currency, thanhTienNha As Currency, thanhTienNhaVaDat As Currency
    
    ' Doc gia tri tu cot Q va R, neu rong thi coi nhu la 0
    thanhTienDat = Nz(wsData.Range("Q" & activeRow).Value, 0)
    thanhTienNha = Nz(wsData.Range("R" & activeRow).Value, 0)
    
    ' Tinh tong va ghi vao cot S
    thanhTienNhaVaDat = thanhTienDat + thanhTienNha
    wsData.Range("S" & activeRow).Value = thanhTienNhaVaDat
    
    ' Neu tong gia tri la 0 thi khong can tinh toan tiep
    If thanhTienNhaVaDat = 0 Then Exit Sub
    
    '--- Tim dong tien do ---
    Dim dongTienDo As Long, timThay As Boolean
    timThay = False
    For dongTienDo = 1 To wsTienDo.Cells(wsTienDo.Rows.Count, "C").End(xlUp).row
        If wsTienDo.Cells(dongTienDo, "C").Value = tenTienDoCanTim Then timThay = True: Exit For
    Next dongTienDo
    If Not timThay Then Exit Sub
    
    '--- Lay va ghi Ti Le TT cua Dot 1 ---
    Dim tiLeDot1 As Variant
    tiLeDot1 = wsTienDo.Cells(dongTienDo, 5).Value 'Cot E la ty le % dot 1
    If IsNumeric(tiLeDot1) And tiLeDot1 <> "" Then
        wsData.Range(colTiLeTT & activeRow).Value = tiLeDot1
    Else
        wsData.Range(colTiLeTT & activeRow).ClearContents
    End If
    
    '--- BUOC 1: Tim dot cuoi cung co gia tri ---
    Dim dotCuoiCung As Integer, i As Integer
    dotCuoiCung = 0
    For i = 16 To 1 Step -1
        If IsNumeric(wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value) And wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value <> "" Then
            dotCuoiCung = i
            Exit For
        End If
    Next i
    
    If dotCuoiCung = 0 Then Exit Sub
    
    '--- BUOC 2: Tinh toan cac dot ---
    Dim tongTienDaTra As Currency, soTienPhaiTra As Currency, tyLePhanTram As Variant
    Dim colIndexOutput As Long, colIndexBC_Dot1 As Long
    Dim colIndexNgayOutput As Long, ngayThanhToanHienTai As Date, ngayTiepTheo As Date, soNgayCongThem As Variant
    
    tongTienDaTra = 0
    colIndexOutput = wsData.Range(colBatDauTraTien & 1).Column
    colIndexBC_Dot1 = wsData.Range(colBC_Dot1 & 1).Column
    colIndexNgayOutput = wsData.Range(colBatDauNgayTT & 1).Column
    If Not IsDate(wsData.Cells(activeRow, colIndexNgayOutput).Value) Then Exit Sub
    ngayThanhToanHienTai = wsData.Cells(activeRow, colIndexNgayOutput).Value
    
    For i = 1 To dotCuoiCung
        If i < dotCuoiCung Then
            tyLePhanTram = wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value
            soTienPhaiTra = VBA.Round(thanhTienNhaVaDat * tyLePhanTram, 0)
            tongTienDaTra = tongTienDaTra + soTienPhaiTra
        Else
            soTienPhaiTra = thanhTienNhaVaDat - tongTienDaTra
        End If
        
        ' Ghi so tien bang so
        wsData.Cells(activeRow, colIndexOutput + (i - 1) * 2).Value = soTienPhaiTra
        
        ' *** GIU NGUYEN: Ghi so tien bang chu, su dung ham vnd() ***
        wsData.Cells(activeRow, colIndexBC_Dot1 + i - 1).Value = vnd(soTienPhaiTra)
        
        ' Tinh toan ngay thang cho cac dot tiep theo
        If i > 1 Then
            soNgayCongThem = wsTienDo.Cells(dongTienDo, (i - 2) * 2 + 6).Value
            If IsNumeric(soNgayCongThem) And soNgayCongThem <> "" Then
                ngayTiepTheo = DateAdd("d", soNgayCongThem, ngayThanhToanHienTai)
                wsData.Cells(activeRow, colIndexNgayOutput + (i - 1) * 2).Value = ngayTiepTheo
                ngayThanhToanHienTai = ngayTiepTheo
            Else
                wsData.Cells(activeRow, colIndexNgayOutput + (i - 1) * 2).ClearContents
            End If
        End If
    Next i
    
    '--- BUOC 3: Xoa du lieu thua o cac dot sau dot cuoi cung ---
    For i = dotCuoiCung + 1 To 16
        wsData.Cells(activeRow, colIndexOutput + (i - 1) * 2).ClearContents
        wsData.Cells(activeRow, colIndexBC_Dot1 + i - 1).ClearContents
        wsData.Cells(activeRow, colIndexNgayOutput + (i - 1) * 2).ClearContents
    Next i
End Sub

'Luu y: Ham Nz (Non-Zero) la mot ham tich hop san trong VBA (Access) nhung khong co san trong Excel.
'Ban can them ham nay vao module de tranh loi.
Public Function Nz(ByVal Value As Variant, Optional ByVal ValueIfNull As Variant = 0) As Variant
    If IsNull(Value) Or IsEmpty(Value) Or Value = "" Then
        Nz = ValueIfNull
    Else
        Nz = Value
    End If
End Function
