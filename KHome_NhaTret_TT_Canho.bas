Attribute VB_Name = "KHome_NhaTret_TT_Canho"
Option Explicit
Sub TinhToanTongHop_ChoDongHienTai()
    '--- KHAI BAO ---
    Dim wsSetup As Worksheet, wsData As Worksheet
    Dim row As Range
    Dim activeRow As Long
    
    '--- KHAI BAO CAC BIEN LUU TRU DONG BI LOI ---
    Dim skippedRows As String
    Dim processedCount As Long
    skippedRows = ""
    processedCount = 0
    
    '--- KHAI BAO CAU HINH ---
    Dim colTienDat As String, colDtThongThuy As String, colTenTienDo As String, colBatDauNgayTT As String
    Dim colTienNha As String, colGiaTriQSDD As String, ColNhaVaDat As String, colPhiBaoTri As String
    Dim colBC_NhaVaDat As String, colBC_GiaTriCH As String, colBC_TienDatCoc As String, colBC_ThueGTGT As String, colBC_BatDauDot1 As String

    '--- DOC CAU HINH ---
    On Error Resume Next
    Set wsSetup = ThisWorkbook.Sheets("Setup")
    If wsSetup Is Nothing Then
        MsgBox "Khong tim thay sheet 'Setup'.", vbCritical, "Loi cau hinh"
        Exit Sub
    End If
    On Error GoTo 0
    
    With wsSetup
        colTienDat = .Range("B1").Value
        colTienNha = .Range("B2").Value
        ColNhaVaDat = .Range("B3").Value
        colTenTienDo = .Range("B4").Value
        colBatDauNgayTT = .Range("B6").Value
        colBC_NhaVaDat = .Range("B7").Value
        colBC_TienDatCoc = .Range("B8").Value
        colBC_BatDauDot1 = .Range("B9").Value
    End With

    '--- KHOI TAO ---
    Set wsData = ThisWorkbook.Sheets("FILE TONG HOA PHU - K HOME")
    Application.ScreenUpdating = False

    '*** VONG LAP QUA TUNG DONG VA KIEM TRA XEM DONG CO BI AN KHONG ***
    For Each row In Selection.Rows
        If row.Hidden = False Then
            activeRow = row.row
            
            '========================================================================
            '   *** KIEM TRA DIEU KIEN TRUOC KHI XU LY TUNG DONG ***
            '========================================================================
            If wsData.Range(colTenTienDo & activeRow).Value = "" Then
                skippedRows = skippedRows & "Dong " & activeRow & ": Thieu Tien Do" & vbCrLf
            ElseIf Not IsDate(wsData.Range(colBatDauNgayTT & activeRow).Value) Then
                skippedRows = skippedRows & "Dong " & activeRow & ": Thieu hoac sai Ngay TT Dot 1" & vbCrLf
            Else
                processedCount = processedCount + 1
                
                '========================================================================
                '   PHAN 1: GOI SUB TAO SO HOP DONG
                '========================================================================
                Call TaoSoHopDong(activeRow)

                '========================================================================
                '   PHAN 2: TINH TOAN CAC GIA TRI CO BAN CUA CAN HO
                '========================================================================
                Call TinhTienDoThanhToan(activeRow)
            End If
        End If
    Next row
    
    Application.ScreenUpdating = True
    
    '========================================================================
    '   HIEN THI THONG BAO TONG KET CUOI CUNG
    '========================================================================
    Dim finalMsg As String
    finalMsg = "Hoan tat!" & vbCrLf & vbCrLf
    finalMsg = finalMsg & "So dong da xu ly thanh cong: " & processedCount & vbCrLf & vbCrLf
    
    If skippedRows <> "" Then
        finalMsg = finalMsg & "Cac dong sau da bi bo qua:" & vbCrLf & skippedRows
    End If
    
    MsgBox finalMsg, vbInformation, "Ket qua tinh toan"
End Sub

