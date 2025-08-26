Option Explicit

'--------------------------------------------------------------------------------------------------
Function LayDuongDanHopLe(ByVal wsNguon As Worksheet) As String
    Dim folderPath As String
    Dim fso As Object
    Dim fldrPicker As FileDialog
    
    ' Doc duong dan tu o P2 cua sheet nguon
    folderPath = wsNguon.Range("Q2").Value
    
    ' Tao doi tuong de kiem tra thu muc
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Kiem tra xem duong dan co ton tai khong
    If folderPath = "" Or Not fso.FolderExists(folderPath) Then
        ' Neu khong, hien cua so cho nguoi dung chon
        MsgBox "Duong dan luu file trong o Q2 khong hop le hoac bi trong." & vbCrLf & _
               "Vui long chon mot thu muc de luu file.", vbInformation, "Thong Bao"
               
        Set fldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
        fldrPicker.Title = "Vui long chon thu muc de luu bao cao"
        
        If fldrPicker.Show = True Then
            ' Neu nguoi dung chon, lay duong dan va luu vao o P2
            folderPath = fldrPicker.SelectedItems(1)
            wsNguon.Range("Q2").Value = folderPath
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
    Dim i As Long
    Dim sourceLastRow As Long
    Dim fullSavePath As String, folderPath As String, fileName As String, fileNamePart As String

    ' --- THIET LAP BAN DAU & XU LY LOI ---
    Set sourceSheet = Sheet3 ' >>> Ban hay kiem tra lai Sheet3 co dung la sheet nguon khong.
    Dim dotSo As String
    dotSo = sourceSheet.Range("P2").Value
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
    ' Vung du lieu nguon bao gom ca cot STT (A)
    Set sourceHeaderRange = sourceSheet.Range("A2:R2")
    If sourceLastRow >= 3 Then
        Set sourceDataRange = sourceSheet.Range("A3:R" & sourceLastRow)
        Set sourceVisibleData = sourceDataRange.SpecialCells(xlCellTypeVisible)
    End If
    
    ' --- TAO FILE MOI VA SAO CHEP DU LIEU ---
    Set newWb = Workbooks.Add
    Set newSheet = newWb.Sheets(1)
    
    ' Sao chep dong tieu de
    sourceHeaderRange.Copy
    With newSheet.Range("A1")
        .PasteSpecial Paste:=xlPasteColumnWidths
        .PasteSpecial Paste:=xlPasteFormats
        .PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    End With
    Application.CutCopyMode = False
    
    ' Sao chep vung du lieu
    If Not sourceVisibleData Is Nothing Then
        sourceVisibleData.Copy
        ' === SUA LOI: Dan du lieu vao cot A, vi du lieu goc da co STT ===
        newSheet.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    End If
    
    ' --- DINH DANG VA XU LY FILE MOI ---
    With newSheet
        .Cells.EntireRow.AutoFit
        
        ' === XOA BO PHAN TU TAO STT VI KHONG CAN THIET ===
        
        ' Dinh nghia va dinh dang cho o chua ngay thang
        Dim finalDataRow As Long
        finalDataRow = .Cells(.Rows.Count, "B").End(xlUp).row
        With .Cells(finalDataRow + 3, "G")
            .Value = Date
            .NumberFormat = "dd/MM/yyyy"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        .Range("P:R").ClearContents
        
        ' Dinh dang dong tieu de A1:N1
        With .Range("A1:O1")
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenterAcrossSelection
        End With
        ' Dinh dang dong du lieu dau tien A2:N2
        With .Range("A2:O2")
            .VerticalAlignment = xlCenter    ' Canh giua theo chieu doc (Middle)
            .HorizontalAlignment = xlCenter  ' Canh giua theo chieu ngang (Center)
            .WrapText = True                 ' Xuong dong tu dong
        End With

        ' Ke vien cho toan bo bang
        Set borderRange = .Range("A1:O" & finalDataRow + 3)
        With borderRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        borderRange.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium

        If .DrawingObjects.Count > 0 Then .DrawingObjects.Delete
        
        ' Doi ten sheet va lam sach ten file
        If dotSo <> "" Then
            .Name = dotSo
            fileNamePart = dotSo
        Else
            .Name = "BaoCao" ' Ten mac dinh neu P2 trong
            fileNamePart = "BaoCao"
        End If
    End With
    
   ' --- LUU FILE MOI (Phien ban chong trung ten file) ---
    
    ' Lam sach ten file de tranh loi
    fileNamePart = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(fileNamePart, "/", "-"), "\", "-"), ":", ""), "*", ""), "?", ""), """", ""), "<", ""), ">", ""), "|", "")
    
    ' Dam bao duong dan luon co dau "\" o cuoi
    If Right(folderPath, 1) <> Application.PathSeparator Then folderPath = folderPath & Application.PathSeparator
    
    ' Kiem tra va tao ten file duy nhat
    Dim baseName As String
    Dim fileExtension As String
    Dim counter As Long
    
    baseName = folderPath & "K-HOME CAN HO_Ð_" & fileNamePart & "_" & Format(Date, "yyyymmdd")
    fileExtension = ".xlsx"
    counter = 0
    
    fullSavePath = baseName & fileExtension
    
    ' Vong lap de kiem tra neu file da ton tai
    While Dir(fullSavePath) <> ""
        counter = counter + 1
        fullSavePath = baseName & " (" & counter & ")" & fileExtension
    Wend
    
    ' Luu file voi ten da duoc dam bao la duy nhat
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


