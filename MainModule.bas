Option Explicit

'======================================================='
'   MODULE CHINH: MainModule                            '
'======================================================='

Private Const SHEET_SETUP = "Setup"
Private Const SHEET_DATA = "FILE TONG HOA PHU - K HOME"
Private Const SHEET_TIENDO = "TIEN_DO_TT"
Private Const MAX_PAYMENT_PERIODS As Integer = 20

Sub TinhToanTongHop_NhaTret_Final()
    Dim wsData As Worksheet, wsTienDo As Worksheet, row As Range
    Dim nhaTret As clsNhaTret, config As Object
    Dim skippedRows As String, processedCount As Long
    skippedRows = "": processedCount = 0
    
    On Error GoTo InitializeError
    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)
    Set wsTienDo = ThisWorkbook.Sheets(SHEET_TIENDO)
    Set config = ReadConfig(SHEET_SETUP)
    If config Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ProcessError

    For Each row In Selection.Rows
        If Not row.Hidden Then
            Set nhaTret = New clsNhaTret
            LoadNhaTretData nhaTret, wsData, config, row.row
            
            Dim totalPct As Double
            totalPct = SumSchedulePercentages(wsTienDo, nhaTret.TenTienDo)
            
            ' (tu? ch?n) c?nh báo n?u t?ng % > 1
            'If totalPct > 1 + 0.000001 Then Debug.Print "WARNING % > 100% at row "; row.Row
            
            nhaTret.XacDinhGiaTriGoc totalPct
            nhaTret.TinhTienDoThanhToan
            nhaTret.TaoSoHopDong
            WriteResultsToSheet nhaTret, wsData, config
            
            processedCount = processedCount + 1
        End If
    Next row

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ShowSummaryMsg processedCount, skippedRows
    Exit Sub

InitializeError:
    MsgBox "Loi khoi tao: " & Err.Description, vbCritical, "Loi He Thong"
    GoTo CleanUp
ProcessError:
    skippedRows = skippedRows & "Dong " & row.row & ": Loi xu ly - " & Err.Description & vbCrLf
    Resume Next
End Sub

Private Function ReadConfig(ByVal setupSheetName As String) As Object
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(setupSheetName)
    Set ReadConfig = CreateObject("Scripting.Dictionary")
    
    With ReadConfig
        .Add "colThanhTienDat_Input", Trim(ws.Range("B1").Value)   ' Q
        .Add "colThanhTienNha_Input", Trim(ws.Range("B2").Value)   ' R
        .Add "colThanhTien", Trim(ws.Range("B3").Value)            ' S
        .Add "colTenTienDo", Trim(ws.Range("B4").Value)            ' T
        .Add "colStartTienTT", Trim(ws.Range("B5").Value)          ' AB
        .Add "colNgayTT1", Trim(ws.Range("B6").Value)              ' AA
        .Add "colBC_ThanhTien", Trim(ws.Range("B7").Value)         ' BE
        .Add "colBC_TienCoc", Trim(ws.Range("B8").Value)           ' BF
        .Add "colStartBC", Trim(ws.Range("B9").Value)              ' BG (BC_dot_1)
        .Add "colCoc_NonHDMB_Output", Trim(ws.Range("B10").Value)  ' F (ví d?)
        .Add "colLoO", Trim(ws.Range("B11").Value)                 ' F
        .Add "colNgayKy", Trim(ws.Range("B12").Value)              ' H
        .Add "colSoHD", Trim(ws.Range("B13").Value)                ' I
        .Add "colTiLeTT_Dot1_Output", Trim(ws.Range("B14").Value)  ' V
        .Add "colKiemTra", Trim(ws.Range("B15").Value)               ' U
        .Add "colBC_ThanhTien_Dat", Trim(ws.Range("B16").Value)      ' BV  <<< M?I
        .Add "colBC_ThanhTien_Nha", Trim(ws.Range("B17").Value)      ' BW  <<< M?I
    End With
    Exit Function
ErrorHandler:
    MsgBox "Loi doc cau hinh tu sheet '" & setupSheetName & "': " & Err.Description, vbCritical
    Set ReadConfig = Nothing
End Function

Private Sub LoadNhaTretData(ByVal nhaTret As clsNhaTret, ByVal ws As Worksheet, ByVal conf As Object, ByVal r As Long)
    With nhaTret
        .RowNum = r
        .TongThanhTien = ws.Range(conf("colThanhTien") & r).Value
        .MaSoLo = ws.Range(conf("colLoO") & r).Value
        .NgayKy = ws.Range(conf("colNgayKy") & r).Value
        .TenTienDo = Trim(CStr(ws.Range(conf("colTenTienDo") & r).Value))
        .NgayTTDot1 = ws.Range(conf("colNgayTT1") & r).Value
        .ThanhTienDat_Input = ws.Range(conf("colThanhTienDat_Input") & r).Value
        .ThanhTienNha_Input = ws.Range(conf("colThanhTienNha_Input") & r).Value
    End With
End Sub

Private Function SumSchedulePercentages(ByVal wsTienDo As Worksheet, ByVal scheduleName As String) As Double
    ' Tr? v? t?ng % theo D?NG TH?P PHÂN (ví d? 0.3 = 30%)
    Dim total As Double: total = 0
    If Len(Trim(scheduleName)) = 0 Then Exit Function
    Dim r As Long, i As Integer, lastRow As Long
    lastRow = wsTienDo.Cells(wsTienDo.Rows.Count, "C").End(xlUp).row
    For r = 1 To lastRow
        If Trim$(UCase$(wsTienDo.Cells(r, "C").Value)) = Trim$(UCase$(scheduleName)) Then
            For i = 5 To (5 + (MAX_PAYMENT_PERIODS - 1) * 2) Step 2
                Dim val As Variant: val = wsTienDo.Cells(r, i).Value
                If IsNumeric(val) And Len(CStr(val)) > 0 Then total = total + CDbl(val)
            Next i
            SumSchedulePercentages = total
            Exit Function
        End If
    Next r
End Function

Private Sub WriteResultsToSheet(ByVal nhaTret As clsNhaTret, ByVal ws As Worksheet, ByVal conf As Object)
    Dim r As Long: r = nhaTret.RowNum
    With ws
        ' S? HÐ
        .Range(conf("colSoHD") & r).Value = nhaTret.SoHopDong
        
        ' C?t m?c
        Dim colTien As Long, colNgay As Long, colBC As Long
        colTien = .Range(conf("colStartTienTT") & 1).Column
        colNgay = .Range(conf("colNgayTT1") & 1).Column
        colBC = .Range(conf("colStartBC") & 1).Column
        
        Dim bcThanhTienCol As Long
        bcThanhTienCol = .Range(conf("colBC_ThanhTien") & 1).Column
        
        ' Hai c?t m?i
        Dim bcThanhTienDatCol As Long, bcThanhTienNhaCol As Long
        bcThanhTienDatCol = .Range(conf("colBC_ThanhTien_Dat") & 1).Column
        bcThanhTienNhaCol = .Range(conf("colBC_ThanhTien_Nha") & 1).Column

        ' --- CLEAR: b? qua BE, BV, BW ---
        Dim i As Integer, tgtCol As Long
        For i = 1 To MAX_PAYMENT_PERIODS
            .Cells(r, colTien + (i - 1) * 2).ClearContents
            .Cells(r, colNgay + (i - 1) * 2).ClearContents
            tgtCol = colBC + i - 1
            If tgtCol <> bcThanhTienCol _
               And tgtCol <> bcThanhTienDatCol _
               And tgtCol <> bcThanhTienNhaCol Then
                .Cells(r, tgtCol).ClearContents
            End If
        Next i
        
        ' T? l? d?t 1
        .Range(conf("colTiLeTT_Dot1_Output") & r).Value = nhaTret.TiLeThanhToanDot1
        
        ' C?c và b?ng ch? ti?n c?c
        If Not nhaTret.IsHDMBContract Then
            ' %_TI?N_C?C = THÀNH_TI?N × t?ng %
            Dim tienCocVal As Currency
            tienCocVal = nhaTret.TongThanhTien * SumSchedulePercentages(ThisWorkbook.Sheets(SHEET_TIENDO), nhaTret.TenTienDo)
            
            .Range(conf("colCoc_NonHDMB_Output") & r).Value = tienCocVal
            
            If tienCocVal > 0 Then
                .Range(conf("colBC_TienCoc") & r).Value = vnd(tienCocVal)
            Else
                .Range(conf("colBC_TienCoc") & r).ClearContents
            End If
        Else
            .Range(conf("colCoc_NonHDMB_Output") & r).ClearContents
            .Range(conf("colBC_TienCoc") & r).ClearContents
        End If

        
        ' L?ch thanh toán & BC_Ð?t_i
        Dim scheduleArray As Variant
        scheduleArray = nhaTret.TienDoThanhToan
        
        Dim sumKiemTra As Currency: sumKiemTra = 0
        
        If IsArray(scheduleArray) Then
            On Error Resume Next
            Dim ub As Long: ub = UBound(scheduleArray, 1)
            On Error GoTo 0
            If ub > 0 Then
                For i = 1 To ub
                    Dim soTien As Currency: soTien = scheduleArray(i, 1)
                    .Cells(r, colTien + (i - 1) * 2).Value = soTien
                    tgtCol = colBC + i - 1
                    If tgtCol <> bcThanhTienCol Then
                        .Cells(r, tgtCol).Value = vnd(soTien)
                    End If
                    If IsDate(scheduleArray(i, 2)) Then
                        .Cells(r, colNgay + (i - 1) * 2).Value = CDate(scheduleArray(i, 2))
                    End If
                    If IsNumeric(soTien) Then sumKiemTra = sumKiemTra + soTien
                Next i
            End If
        End If
        
        ' KI?M_TRA (Setup!B15)
        Dim colKiemTra As Long
        colKiemTra = .Range(conf("colKiemTra") & 1).Column
        .Cells(r, colKiemTra).Value = sumKiemTra
        ' N?u mu?n b?ng ch?: .Cells(r, colKiemTra).Value = vnd(sumKiemTra)
        ' === Ghi BC_THÀNH_TI?N_Ð?T (BV) ===
        If IsNumeric(nhaTret.ThanhTienDat_Input) And nhaTret.ThanhTienDat_Input > 0 Then
            .Cells(r, bcThanhTienDatCol).Value = vnd(nhaTret.ThanhTienDat_Input)
        Else
            .Cells(r, bcThanhTienDatCol).ClearContents
        End If

        ' === Ghi BC_THÀNH_TI?N_NHÀ (BW) ===
        If IsNumeric(nhaTret.ThanhTienNha_Input) And nhaTret.ThanhTienNha_Input > 0 Then
            .Cells(r, bcThanhTienNhaCol).Value = vnd(nhaTret.ThanhTienNha_Input)
        Else
            .Cells(r, bcThanhTienNhaCol).ClearContents
        End If
        ' BC_THÀNH_TI?N (ghi cu?i)
        If IsNumeric(nhaTret.TongThanhTien) And nhaTret.TongThanhTien > 0 Then
            .Range(conf("colBC_ThanhTien") & r).Value = vnd(nhaTret.TongThanhTien)
        Else
            .Range(conf("colBC_ThanhTien") & r).ClearContents
        End If
    End With
End Sub

Private Sub ShowSummaryMsg(ByVal processedCount As Long, ByVal skippedRows As String)
    Dim finalMsg As String
    finalMsg = "Hoan tat!" & vbCrLf & vbCrLf & _
               "So dong da xu ly thanh cong: " & processedCount & vbCrLf & vbCrLf
    If Len(skippedRows) > 0 Then
        finalMsg = finalMsg & "Cac dong sau da bi bo qua:" & vbCrLf & skippedRows
    End If
    MsgBox finalMsg, vbInformation, "Ket qua tinh toan"
End Sub


