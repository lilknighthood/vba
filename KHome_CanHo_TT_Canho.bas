Option Explicit
Sub TinhToanTongHop_ChoDongHienTai()
    '--- KHAI BAO ---
    Dim wsSetup As Worksheet, wsData As Worksheet
    Dim row As Range, activeRow As Long
    Dim skippedRows As String, processedCount As Long
    skippedRows = "": processedCount = 0
    
    '--- KHAI BAO CAU HINH ---
    Dim colGiaBan As String, colDtThongThuy As String, colTenTienDo As String, colBatDauNgayTT As String
    Dim colGiaTriCanHo As String, colGiaTriQSDD As String, colThueGTGT As String, colPhiBaoTri As String
    Dim colBC_GiaBan As String, colBC_GiaTriCH As String, colBC_GiaTriQSDD As String, colBC_ThueGTGT As String, colBC_PhiBaoTri As String

    '--- DOC CAU HINH ---
    On Error Resume Next
    Set wsSetup = ThisWorkbook.Sheets("Setup")
    If wsSetup Is Nothing Then MsgBox "Khong tim thay sheet 'Setup'.", vbCritical, "Loi cau hinh": Exit Sub
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

    '--- KHOI TAO ---
    Set wsData = ThisWorkbook.Sheets("CAN HO K-HOME")
    Application.ScreenUpdating = False

    For Each row In Selection.Rows
        If row.Hidden = False Then
            activeRow = row.row
            
            If wsData.Range(colTenTienDo & activeRow).Value = "" Then
                skippedRows = skippedRows & "Dong " & activeRow & ": Thieu Tien Do" & vbCrLf
            ElseIf Not IsDate(wsData.Range(colBatDauNgayTT & activeRow).Value) Then
                skippedRows = skippedRows & "Dong " & activeRow & ": Thieu hoac sai Ngay TT Dot 1" & vbCrLf
            Else
                processedCount = processedCount + 1
                Call TaoSoHopDong(activeRow)

                Dim heSoDat As Double, giaBanCanHo As Currency, dtThongThuy As Double
                Dim giaTriQSDD As Currency, giaTriCanHo As Currency, thueGTGT As Currency, phiBaoTri As Currency
                
                heSoDat = 729754.9204
                giaBanCanHo = wsData.Range(colGiaBan & activeRow).Value
                dtThongThuy = wsData.Range(colDtThongThuy & activeRow).Value
                
                If giaBanCanHo > 0 And dtThongThuy > 0 Then
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
                    
                    'Goi sub phu duy nhat, truyen cac gia tri can thiet
                    Call TinhTienDoThanhToan(activeRow, giaBanCanHo, giaTriCanHo)
                Else
                     skippedRows = skippedRows & "Dong " & activeRow & ": Loi du lieu (Gia ban hoac DTSD)" & vbCrLf
                End If
            End If
        End If
    Next row
    
    Application.ScreenUpdating = True
    
    Dim finalMsg As String
    finalMsg = "Hoan tat!" & vbCrLf & vbCrLf
    finalMsg = finalMsg & "So dong da xu ly thanh cong: " & processedCount & vbCrLf & vbCrLf
    
    If skippedRows <> "" Then
        finalMsg = finalMsg & "Cac dong sau da bi bo qua:" & vbCrLf & skippedRows
    End If
    
    MsgBox finalMsg, vbInformation, "Ket qua tinh toan"
End Sub
