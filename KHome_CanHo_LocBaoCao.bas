'==================================================================================================
'   Muc dich: Loc bao cao, ghi STT, ghi ten nguoi lap va tu dong ke vien cho bang du lieu.
'==================================================================================================
Sub LocBaoCao()
    ' --- KHAI BAO CAC BIEN VA HANG SO ---
    Const SOURCE_SHEET_NAME As String = "CAN HO K-HOME" ' <<< Ten sheet du lieu nguon
    Const REPORT_SHEET_NAME As String = "TRINH_KY"      ' <<< Ten sheet trinh ky (bao cao)
    Const INPUT_CELL As String = "F1"                   ' <<< O chua dieu kien loc "DOT THANH TOAN"
    Const ALL_ITEMS_TEXT As String = "Tat Ca"           ' <<< Text de chon tat ca
    
    ' Khai bao cac bien
    Dim sourceSheet As Worksheet, reportSheet As Worksheet
    Dim sourceData As Variant, resultsData() As Variant, sttArray() As Variant
    Dim filterCondition As String
    Dim lastRow As Long, resultCounter As Long
    Dim i As Long
    
    ' --- TANG TOC DO THUC THI & XU LY LOI ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo CleanupAndExit

    ' --- THIET LAP CAC SHEET LAM VIEC ---
    Set sourceSheet = ThisWorkbook.Sheets(SOURCE_SHEET_NAME)
    Set reportSheet = ThisWorkbook.Sheets(REPORT_SHEET_NAME)
    filterCondition = LCase(reportSheet.Range(INPUT_CELL).Value)

    ' --- DOC DU LIEU TU SHEET NGUON VAO MANG ---
    With sourceSheet
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).row
        If lastRow < 2 Then
            MsgBox "Khong co du lieu tai sheet '" & SOURCE_SHEET_NAME & "'.", vbInformation, "Thong bao"
            GoTo CleanupAndExit
        End If
        sourceData = .Range("B2:DC" & lastRow).Value
    End With

    ' --- XU LY DU LIEU ---
    ReDim resultsData(1 To UBound(sourceData, 1), 1 To 17)
    resultCounter = 0
    Const COL_DOT_TT As Long = 29
    
    For i = 1 To UBound(sourceData, 1)
        If filterCondition = LCase(ALL_ITEMS_TEXT) Or LCase(CStr(sourceData(i, COL_DOT_TT))) = filterCondition Then
            resultCounter = resultCounter + 1
            resultsData(resultCounter, 1) = sourceData(i, 4)
            resultsData(resultCounter, 2) = sourceData(i, 8)
            resultsData(resultCounter, 3) = sourceData(i, 9)
            resultsData(resultCounter, 4) = sourceData(i, 10)
            resultsData(resultCounter, 5) = sourceData(i, 11)
            resultsData(resultCounter, 6) = sourceData(i, 12)
            resultsData(resultCounter, 7) = sourceData(i, 20)
            resultsData(resultCounter, 8) = sourceData(i, 28)
            resultsData(resultCounter, 9) = sourceData(i, 26)
            resultsData(resultCounter, 10) = sourceData(i, 33)
            resultsData(resultCounter, 11) = sourceData(i, 34)
            resultsData(resultCounter, 12) = sourceData(i, 35)
            resultsData(resultCounter, 13) = sourceData(i, 31)
        End If
    Next i

    ' --- GHI KET QUA RA SHEET TRINH KY ---
    With reportSheet
        ' Xoa du lieu cu va vien cu
        .Range("A4:N" & .Rows.Count).ClearContents
        .Range("A3:N" & .Rows.Count).Borders.LineStyle = xlNone
        
        If resultCounter > 0 Then
            ' === PHAN MOI THEM: TAO SO THU TU CHO COT A ===
            ReDim sttArray(1 To resultCounter, 1 To 1)
            For i = 1 To resultCounter
                sttArray(i, 1) = i
            Next i
            ' Ghi toan bo mang STT vao cot A, bat dau tu A4
            .Range("A4").Resize(resultCounter, 1).Value = sttArray
            ' === KET THUC PHAN TAO SO THU TU ===
            
            ' Ghi du lieu chinh vao B4:N
            .Range("B4").Resize(resultCounter, 13).Value = resultsData
            
            ' Ghi ten nguoi lap
            Dim lastDataRow As Long
            lastDataRow = 4 + resultCounter - 1
            .Cells(lastDataRow + 2, "G").Value = "NGUY" & ChrW(7876) & "N TH" & ChrW(7882) & " M" _
        & ChrW(7896) & "NG TUY" & ChrW(7870) & "T"
            
            ' Ke vien cho toan bo bang du lieu
            Dim borderRange As Range
            Set borderRange = .Range("A3:N" & lastDataRow + 2)
            With borderRange.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            borderRange.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            
        Else
            ' Thong bao neu khong tim thay du lieu
            MsgBox "KHONG TIM THAY DU LIEU PHU HOP VOI DOT: '" & .Range(INPUT_CELL).Value & "'", vbExclamation, "Thong bao"
        End If
    End With

' --- DON DEP VA KET THUC ---
CleanupAndExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Set sourceSheet = Nothing
    Set reportSheet = Nothing
    
    If Err.Number <> 0 Then
        MsgBox "Da co loi xay ra!" & vbCrLf & vbCrLf & _
               "Ma loi: " & Err.Number & vbCrLf & _
               "Mo ta: " & Err.Description, _
               vbCritical, "Loi Thuc Thi Macro"
    End If
End Sub

