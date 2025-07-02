Option Explicit
Sub TaoSoHopDong(ByVal activeRow As Long)
    '--- KHAI BAO ---
    Dim wsSetup As Worksheet, wsData As Worksheet
    
    '--- KHAI BAO CAU HINH ---
    Dim colLoO As String
    Dim colNgayKy As String
    Dim colSoHD As String

    '--- DOC CAU HINH ---
    On Error Resume Next
    Set wsSetup = ThisWorkbook.Sheets("Setup")
    If wsSetup Is Nothing Then Exit Sub 'Thoat neu khong co sheet Setup
    
    With wsSetup
        colLoO = .Range("B11").Value
        colNgayKy = .Range("B12").Value
        colSoHD = .Range("B13").Value
    End With
    
    If colLoO = "" Or colNgayKy = "" Or colSoHD = "" Then Exit Sub 'Thoat neu thieu cau hinh

    '--- KHOI TAO ---
    Set wsData = ThisWorkbook.Sheets("FILE TONG HOA PHU - K HOME")

    '--- Lay du lieu dau vao ---
    Dim maLoO As String
    Dim ngayKy As Variant
    
    maLoO = wsData.Range(colLoO & activeRow).Value
    ngayKy = wsData.Range(colNgayKy & activeRow).Value
    
    '--- Kiem tra dieu kien truoc khi tao so hop dong ---
    If maLoO <> "" And IsDate(ngayKy) Then
        'Tao chuoi so hop dong theo dung dinh dang
        wsData.Range(colSoHD & activeRow).Value = maLoO & "/" & Year(ngayKy) & "/H�/NOXH - HP"
    End If
End Sub


