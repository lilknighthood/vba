Option Explicit ' Buoc trinh bien dich phai khai bao tat ca cac bien

'==================================================================================================
'   Muc dich: Loc bao cao, ghi STT, ghi ten, ke vien va luu ra file Excel moi.
'   Phien ban: Da gop va toi uu toan bo code, bo Wrap Text.
'   Tac gia: Gemini
'   Ngay cap nhat: 04/07/2025
'==================================================================================================

'--------------------------------------------------------------------------------------------------
'   HAM CHUC NANG: Kiem tra va lay duong dan hop le tu sheet nguon.
'   Neu duong dan khong co hoac khong hop le, se hien cua so de nguoi dung chon.
'--------------------------------------------------------------------------------------------------
Function LayDuongDanHopLe(ByVal wsNguon As Worksheet) As String
    Dim folderPath As String
    Dim fso As Object
    Dim fldrPicker As FileDialog
    
    ' Doc duong dan tu o P2 cua sheet nguon
    folderPath = wsNguon.Range("P2").Value
    
    ' Tao doi tuong de kiem tra thu muc
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Kiem tra xem duong dan co ton tai khong
    If folderPath = "" Or Not fso.FolderExists(folderPath) Then
        ' Neu khong, hien cua so cho nguoi dung chon
        MsgBox "Duong dan luu file trong o P2 khong hop le hoac bi trong." & vbCrLf & _
               "Vui long chon mot thu muc de luu file.", vbInformation, "Thong Bao"
               
        Set fldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
        fldrPicker.Title = "Vui long chon thu muc de luu bao cao"
        
        If fldrPicker.Show = True Then
            ' Neu nguoi dung chon, lay duong dan va luu vao o P2
            folderPath = fldrPicker.SelectedItems(1)
            wsNguon.Range("P2").Value = folderPath
            LayDuongDanHopLe = folderPath ' Tra ve duong dan da chon
        Else
            ' Neu nguoi dung huy, tra ve chuoi rong
            LayDuongDanHopLe = ""
        End If
    Else
        ' Neu duong dan hop le, tra ve chinh no
        LayDuongDanHopLe = folderPath
    End If
    
    ' Giai phong bo nho
    Set fso = Nothing
    Set fldrPicker = Nothing
End Function


'--------------------------------------------------------------------------------------------------
'   SUB CHINH: Thuc thi toan bo qua trinh tao file bao cao.
'--------------------------------------------------------------------------------------------------
Sub TaoFileBaoCao()
    ' --- KHAI BAO BIEN ---
    Dim sourceSheet As Worksheet, newWb As Workbook, newSheet As Worksheet
    Dim sourceHeaderRange As Range, sourceDataRange As Range, sourceVisibleData As Range, borderRange As Range
    Dim sttArray() As Variant
    Dim sourceLastRow As Long, i As Long
    Dim fullSavePath As String, folderPath As String, fileName As String, fileNamePart As String
    
    ' --- THIET LAP BAN DAU & XU LY LOI ---
    Set sourceSheet = Sheet3 ' >>> Ban hay kiem tra lai Sheet3 co dung la sheet nguon khong.
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    On Error GoTo CleanupAndExit

    ' --- LUU FILE MOI (Thuc hien kiem tra duong dan dau tien) ---
    folderPath = LayDuongDanHopLe(sourceSheet)
    
    ' Neu nguoi dung huy viec chon folder, thoat macro
    If folderPath = "" Then
        MsgBox "Da huy thao tac tao file do khong chon duong dan luu.", vbExclamation
        GoTo CleanupAndExit
    End If
    
    ' --- LAY DU LIEU TU SHEET NGUON ---
    sourceLastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "G").End(xlUp).row
    If sourceLastRow < 2 Then
        MsgBox "Khong co du lieu de tao file.", vbInformation, "Thong bao"
        GoTo CleanupAndExit
    End If
    
    ' --- TACH BIET DONG TIEU DE VA VUNG DU LIEU ---
    Set sourceHeaderRange = sourceSheet.Range("A2:Q2")
    If sourceLastRow >= 3 Then
        Set sourceDataRange = sourceSheet.Range("A3:Q" & sourceLastRow)
        Set sourceVisibleData = sourceDataRange.SpecialCells(xlCellTypeVisible)
    End If
    
    ' --- TAO FILE MOI VA SAO CHEP DU LIEU ---
    Set newWb = Workbooks.Add
    Set newSheet = newWb.Sheets(1)
    
    ' Sao chep dong tieu de voi day du dinh dang
    sourceHeaderRange.Copy
    With newSheet.Range("A1")
        .PasteSpecial Paste:=xlPasteColumnWidths
        .PasteSpecial Paste:=xlPasteAll
    End With
    
    ' Sao chep du lieu va chi dan gia tri
    If Not sourceVisibleData Is Nothing Then
        sourceVisibleData.Copy
        newSheet.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    End If
    
    ' --- DINH DANG VA XU LY FILE MOI ---
    With newSheet
        .Cells.EntireRow.AutoFit ' Tu dong can chinh chieu cao
        
        ' Tao va ghi so thu tu
        Dim dataRowCount As Long
        dataRowCount = .Cells(.Rows.Count, "B").End(xlUp).row - 1
        If dataRowCount > 0 Then
            ReDim sttArray(1 To dataRowCount, 1 To 1)
            For i = 1 To dataRowCount
                sttArray(i, 1) = i
            Next i
            .Range("A3").Resize(dataRowCount, 1).Value = sttArray
        End If
        
        ' Ghi ten nguoi lap va ngay thang
        Dim finalDataRow As Long
        finalDataRow = .Cells(.Rows.Count, "B").End(xlUp).row
        
        ' Dinh nghia va dinh dang cho o chua ngay thang
        With .Cells(finalDataRow + 3, "G")
            .Value = Date                     ' Gan gia tri ngay hien tai
            .NumberFormat = "dd/MM/yyyy"      ' Dinh dang ngay thang
            .Font.Bold = True                 ' To dam
            .HorizontalAlignment = xlCenter   ' Canh giua theo chieu ngang
            .VerticalAlignment = xlCenter     ' Canh giua theo chieu doc
        End With
        
        .Range("P:Q").ClearContents ' Xoa du lieu thua
        With .Range("A2:N2")
            .WrapText = True                 ' Xuong dong tu dong
            .VerticalAlignment = xlCenter    ' Canh giua theo chieu doc (Middle)
            .HorizontalAlignment = xlCenter  ' Canh giua theo chieu ngang (Center)
        End With
        
        ' Ke vien cho toan bo bang
        Set borderRange = .Range("A1:N" & finalDataRow + 3)
        With borderRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        borderRange.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium

        If .DrawingObjects.Count > 0 Then .DrawingObjects.Delete
        
        ' Doi ten sheet va lam sach ten file
        If .Range("O1").Value <> "" Then
            .Name = .Range("O1").Value
            fileNamePart = .Range("O1").Value
        Else
            fileNamePart = "BaoCao" ' Ten mac dinh neu O1 trong
        End If
    End With
    
    ' Lam sach ten file de tranh loi
    fileNamePart = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(fileNamePart, "/", "-"), "\", "-"), ":", ""), "*", ""), "?", ""), """", ""), "<", ""), ">", ""), "|", "")
    fileName = "K-HOME CAN HO Ð_" & fileNamePart & "_" & Format(Date, "yyyymmdd") & ".xlsx"
    
    If Right(folderPath, 1) <> Application.PathSeparator Then folderPath = folderPath & Application.PathSeparator
    
    fullSavePath = folderPath & fileName
    
    newWb.SaveAs fileName:=fullSavePath, FileFormat:=xlOpenXMLWorkbook
    
' --- KET THUC VA DON DEP ---
CleanupAndExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
    
    If Err.Number <> 0 Then
        MsgBox "Da co loi xay ra!" & vbCrLf & vbCrLf & _
               "Ma loi: " & Err.Number & vbCrLf & _
               "Mo ta: " & Err.Description, _
               vbCritical, "Loi Thuc Thi Macro"
    End If
End Sub

