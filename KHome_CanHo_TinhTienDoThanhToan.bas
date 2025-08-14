Option Explicit
'====================================================================================================
' CHUC NANG: Tinh toan chi tiet tien do thanh toan cho mot dong cu the
'            *** PHIEN BAN MOI: Da them tooltip cho cac o Ngay Thanh Toan ***
'====================================================================================================
Sub TinhTienDoThanhToan(ByVal activeRow As Long, ByVal giaBanCanHo As Currency, ByVal giaTriCanHo As Currency)
    '--- KHAI BAO BIEN SO DOT THANH TOAN TOI DA ---
    Const maxSoDotThanhToan As Integer = 20
    
    '--- KHAI BAO DOI TUONG VA BIEN CAU HINH ---
    Dim wsSetup As Worksheet, wsData As Worksheet, wsTienDo As Worksheet
    Dim colTenTienDo As String, colBatDauTraTien As String, colBatDauNgayTT As String
    Dim colTiLeTT As String, colTienCoc As String, colKiemTra As String

    '--- DOC CAU HINH CHO PHAN TIEN DO ---
    Set wsSetup = ThisWorkbook.Sheets("Setup"): Set wsData = ThisWorkbook.Sheets("CAN HO K-HOME"): Set wsTienDo = ThisWorkbook.Sheets("TIEN_DO_TT")
    With wsSetup
        colTenTienDo = .Range("B7").Value: colBatDauTraTien = .Range("B8").Value
        colBatDauNgayTT = .Range("B9").Value
        colTiLeTT = .Range("B16").Value: colTienCoc = .Range("B20").Value
        colKiemTra = .Range("B21").Value
    End With

    '--- TIM DONG TIEN DO TUONG UNG ---
    Dim tenTienDoCanTim As String: tenTienDoCanTim = wsData.Range(colTenTienDo & activeRow).Value
    If tenTienDoCanTim = "" Then Exit Sub
    
    Dim dongTienDo As Long, timThay As Boolean, i As Integer: timThay = False
    For dongTienDo = 1 To wsTienDo.Cells(wsTienDo.Rows.Count, "C").End(xlUp).row
        If wsTienDo.Cells(dongTienDo, "C").Value = tenTienDoCanTim Then
            timThay = True: Exit For
        End If
    Next dongTienDo
    If Not timThay Then Exit Sub
    
    '--- TINH TONG TY LE % VA GHI GIA TRI TIEN COC ---
    Dim tongTyLePhanTram As Double: tongTyLePhanTram = 0
    For i = 1 To maxSoDotThanhToan
        If IsNumeric(wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value) Then
            tongTyLePhanTram = tongTyLePhanTram + wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value
        End If
    Next i
    Dim tienCocValue As Currency
    tienCocValue = giaTriCanHo * tongTyLePhanTram
    wsData.Range(colTienCoc & activeRow).Value = tienCocValue
    
    '--- XAC DINH GIA TRI NEN DE TINH TOAN CAC DOT ---
    Dim baseAmount As Currency
    If UCase(tenTienDoCanTim) Like "*H" & ChrW(272) & "MB*" Then
        baseAmount = giaBanCanHo
    Else
        baseAmount = tienCocValue
    End If
    
    '--- BUOC 1: TIM DOT CUOI CUNG CO GIA TRI % ---
    Dim dotCuoiCung As Integer: dotCuoiCung = 0
    For i = maxSoDotThanhToan To 1 Step -1
        If IsNumeric(wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value) And wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value <> "" Then
            dotCuoiCung = i: Exit For
        End If
    Next i
    
    '--- DON DEP DU LIEU TRUOC KHI TINH ---
    Dim colIndexOutput As Long, colIndexNgayOutput As Long
    colIndexOutput = wsData.Range(colBatDauTraTien & 1).Column
    colIndexNgayOutput = wsData.Range(colBatDauNgayTT & 1).Column
    For i = 1 To maxSoDotThanhToan
        On Error Resume Next
        wsData.Cells(activeRow, colIndexOutput + (i - 1) * 2).Validation.Delete
        wsData.Cells(activeRow, colIndexNgayOutput + (i - 1) * 2).Validation.Delete 'Xoa tooltip cu cua ngay
        On Error GoTo 0
        wsData.Cells(activeRow, colIndexOutput + (i - 1) * 2).ClearContents
        If i > 1 Then wsData.Cells(activeRow, colIndexNgayOutput + (i - 1) * 2).ClearContents
    Next i
    
    If dotCuoiCung = 0 Then
        wsData.Range(colKiemTra & activeRow).ClearContents
        Exit Sub
    End If

    '--- BUOC 2: TINH TOAN CHI TIET CHO TUNG DOT ---
    Dim tongTienDaTra As Currency, soTienPhaiTra As Currency, tyLePhanTram As Variant
    Dim ngayThanhToanHienTai As Date, ngayTiepTheo As Date
    Dim tooltipText As String, targetCell As Range, targetDateCell As Range 'Them bien cho o ngay
    
    tongTienDaTra = 0
    ngayThanhToanHienTai = wsData.Cells(activeRow, colIndexNgayOutput).Value
    
    For i = 1 To dotCuoiCung
        Set targetCell = wsData.Cells(activeRow, colIndexOutput + (i - 1) * 2)
        
        '--- Tinh tien ---
        If i < dotCuoiCung Then
            tyLePhanTram = wsTienDo.Cells(dongTienDo, (i - 1) * 2 + 5).Value
            soTienPhaiTra = VBA.Round(baseAmount * tyLePhanTram, 0)
            tongTienDaTra = tongTienDaTra + soTienPhaiTra
            tooltipText = "Ty le: " & Format(tyLePhanTram, "0.0%") & vbCrLf & "Thanh tien: " & Format(soTienPhaiTra, "#,##0")
        Else
            soTienPhaiTra = baseAmount - tongTienDaTra
            tooltipText = "Phan con lai" & vbCrLf & "Thanh tien: " & Format(soTienPhaiTra, "#,##0")
        End If
        
        targetCell.Value = soTienPhaiTra
        
        'Them tooltip cho o tien
        On Error Resume Next
        targetCell.Validation.Delete
        targetCell.Validation.Add Type:=0, AlertStyle:=1, Operator:=1
        With targetCell.Validation
            .InputTitle = "Chi tiet Dot " & i
            .InputMessage = tooltipText
            .ShowInput = True
            .ShowError = False
        End With
        On Error GoTo 0
        
        '--- Tinh ngay va them tooltip cho ngay (tu dot 2 tro di) ---
        If i > 1 Then
            Set targetDateCell = wsData.Cells(activeRow, colIndexNgayOutput + (i - 1) * 2)
            Dim soNgayCongThem As Variant
            soNgayCongThem = wsTienDo.Cells(dongTienDo, (i - 2) * 2 + 6).Value
            If IsNumeric(soNgayCongThem) And soNgayCongThem <> "" Then
                ngayTiepTheo = DateAdd("d", soNgayCongThem, ngayThanhToanHienTai)
                targetDateCell.Value = ngayTiepTheo
                
                '*** THEM TOOLTIP VAO O NGAY ***
                tooltipText = Format(ngayThanhToanHienTai, "dd/mm/yyyy") & " + " & soNgayCongThem & " ngay"
                On Error Resume Next
                targetDateCell.Validation.Delete
                targetDateCell.Validation.Add Type:=0, AlertStyle:=1, Operator:=1
                With targetDateCell.Validation
                    .InputTitle = "Cong thuc tinh Ngay TT Dot " & i
                    .InputMessage = tooltipText
                    .ShowInput = True
                    .ShowError = False
                End With
                On Error GoTo 0
                
                'Cap nhat ngay hien tai de tinh cho dot tiep theo
                ngayThanhToanHienTai = ngayTiepTheo
            End If
        End If
    Next i
    
    '--- Ghi gia tri cho cot KIEM_TRA ---
    wsData.Range(colKiemTra & activeRow).Value = baseAmount
End Sub
