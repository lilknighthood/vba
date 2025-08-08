Option Explicit

'======================================================='
'   MODULE CHINH: MainModule - Co them cot KIEM_TRA     '
'======================================================='

' Hang so de de bao tri
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
        ' Su dung Trim() de loai bo dau cach thua trong sheet Setup
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
        .Add "colKiemTra", Trim(ws.Range("B15").Value)             ' U
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
    Dim total As Double: total = 0
    If Len(Trim(scheduleName)) = 0 Then Exit Function
    Dim dongTienDo As Long, i As Integer, lastRow As Long
    lastRow = wsTienDo.Cells(wsTienDo.Rows.Count, "C").End(xlUp).row
    For dongTienDo = 1 To lastRow
        If Trim(UCase(wsTienDo.Cells(dongTienDo, "C").Value)) = Trim(UCase(scheduleName)) Then
            For i = 5 To (5 + (MAX_PAYMENT_PERIODS - 1) * 2) Step 2
                Dim val As Variant: val = wsTienDo.Cells(dongTienDo, i).Value
                If IsNumeric(val) And Len(CStr(val)) > 0 Then total = total + CDbl(val)
            Next i
            SumSchedulePercentages = total
            Exit Function
        End If
    Next dongTienDo
End Function

'===================== WRITE (dã lo?i tr? BE kh?i CLEAR) =====================
Private Sub WriteResultsToSheet(ByVal nhaTret As clsNhaTret, ByVal ws As Worksheet, ByVal conf As Object)
    Dim r As Long: r = nhaTret.RowNum
    With ws
        ' 1) S? HÐ
        .Range(conf("colSoHD") & r).Value = nhaTret.SoHopDong
        
        ' 2) Tính các c?t m?c
        Dim colTien As Long, colNgay As Long, colBC As Long
        colTien = .Range(conf("colStartTienTT") & 1).Column
        colNgay = .Range(conf("colNgayTT1") & 1).Column
        colBC = .Range(conf("colStartBC") & 1).Column
        
        ' C?t BC_THÀNH_TI?N (ví d? BE) d? lo?i kh?i CLEAR
        Dim bcThanhTienCol As Long
        bcThanhTienCol = .Range(conf("colBC_ThanhTien") & 1).Column
        
        ' 3) Xoá ti?n/ngày/BC_Ð?T_* (B? QUA c?t BE)
        Dim i As Integer, tgtCol As Long
        For i = 1 To MAX_PAYMENT_PERIODS
            .Cells(r, colTien + (i - 1) * 2).ClearContents
            .Cells(r, colNgay + (i - 1) * 2).ClearContents
            
            tgtCol = colBC + i - 1
            If tgtCol <> bcThanhTienCol Then
                .Cells(r, tgtCol).ClearContents
            End If
        Next i
        
        ' 4) T? l? Ð?t 1
        .Range(conf("colTiLeTT_Dot1_Output") & r).Value = nhaTret.TiLeThanhToanDot1
        
        ' 5) C?c (và b?ng ch? ti?n c?c)
        If Not nhaTret.IsHDMBContract Then
            .Range(conf("colCoc_NonHDMB_Output") & r).Value = nhaTret.GiaTriDeGhiVaoCotCoc
            If nhaTret.GiaTriDeGhiVaoCotCoc > 0 Then
                .Range(conf("colBC_TienCoc") & r).Value = vnd(nhaTret.GiaTriDeGhiVaoCotCoc)
            Else
                .Range(conf("colBC_TienCoc") & r).ClearContents
            End If
        Else
            .Range(conf("colCoc_NonHDMB_Output") & r).ClearContents
            .Range(conf("colBC_TienCoc") & r).ClearContents
        End If
        
        ' 6) Ghi l?ch thanh toán theo d?t + BC_Ð?T_i (b?ng ch?)
        Dim scheduleArray As Variant
        scheduleArray = nhaTret.TienDoThanhToan
        
        Dim sumKiemTra As Currency: sumKiemTra = 0
        
        If IsArray(scheduleArray) Then
            On Error Resume Next
            Dim checkBound As Long: checkBound = UBound(scheduleArray, 1)
            On Error GoTo 0
            
            If checkBound > 0 Then
                For i = 1 To checkBound
                    Dim soTien As Currency: soTien = scheduleArray(i, 1)
                    ws.Cells(r, colTien + (i - 1) * 2).Value = soTien
                    tgtCol = colBC + i - 1
                    If tgtCol <> bcThanhTienCol Then
                        ws.Cells(r, tgtCol).Value = vnd(soTien)
                    End If
                    If IsDate(scheduleArray(i, 2)) Then
                        ws.Cells(r, colNgay + (i - 1) * 2).Value = CDate(scheduleArray(i, 2))
                    End If
                    ' C?ng d?n cho KI?M_TRA
                    If IsNumeric(soTien) Then sumKiemTra = sumKiemTra + soTien
                Next i
            End If
        End If
        
        ' 7) KI?M_TRA = T?ng các d?t thanh toán (Setup!B15)
        Dim colKiemTra As Long
        colKiemTra = .Range(conf("colKiemTra") & 1).Column
        .Cells(r, colKiemTra).Value = sumKiemTra
        ' N?u mu?n b?ng ch?: dùng dòng du?i và b? dòng trên
        ' .Cells(r, colKiemTra).Value = vnd(sumKiemTra)
        
        ' 8) Cu?i cùng ghi BC_THÀNH_TI?N (BE) d? ch?c ch?n không b? xoá
        Dim ttNum As Currency, rawVal As Variant, s As String
        rawVal = nhaTret.TongThanhTien
        s = Trim(CStr(rawVal))
        If Len(s) > 0 Then
            s = Replace(s, ".", "")
            s = Replace(s, " ", "")
            If IsNumeric(s) Then ttNum = CCur(s)
        End If
        
        If ttNum > 0 Then
            .Range(conf("colBC_ThanhTien") & r).Value = vnd(ttNum)
        ElseIf IsNumeric(nhaTret.TongThanhTien) And nhaTret.TongThanhTien > 0 Then
            .Range(conf("colBC_ThanhTien") & r).Value = vnd(nhaTret.TongThanhTien)
        Else
            .Range(conf("colBC_ThanhTien") & r).ClearContents
        End If
    End With
End Sub

Private Sub ShowSummaryMsg(ByVal processedCount As Long, ByVal skippedRows As String)
    Dim finalMsg As String
    finalMsg = "Hoan tat!" & vbCrLf & vbCrLf
    finalMsg = finalMsg & "So dong da xu ly thanh cong: " & processedCount & vbCrLf & vbCrLf
    If Len(skippedRows) > 0 Then
        finalMsg = finalMsg & "Cac dong sau da bi bo qua:" & vbCrLf & skippedRows
    End If
    MsgBox finalMsg, vbInformation, "Ket qua tinh toan"
End Sub


