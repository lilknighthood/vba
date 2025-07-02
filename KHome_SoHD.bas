Attribute VB_Name = "KHome_SoHD"
Option Explicit
Sub TaoSoHopDong(ByVal activeRow As Long)
    '--- KHAI BAO ---
    Dim wsSetup As Worksheet, wsData As Worksheet
    
    '--- KHAI BAO CAU HINH ---
    Dim colCanHo As String
    Dim colNgayKy As String
    Dim colSoHD As String

    '--- DOC CAU HINH ---
    On Error Resume Next
    Set wsSetup = ThisWorkbook.Sheets("Setup")
    If wsSetup Is Nothing Then Exit Sub 'Thoat neu khong co sheet Setup
    
    With wsSetup
        colCanHo = .Range("B17").Value
        colNgayKy = .Range("B18").Value
        colSoHD = .Range("B19").Value
    End With
    
    If colCanHo = "" Or colNgayKy = "" Or colSoHD = "" Then Exit Sub 'Thoat neu thieu cau hinh

    '--- KHOI TAO ---
    Set wsData = ThisWorkbook.Sheets("CAN HO K-HOME")

    '--- Lay du lieu dau vao ---
    Dim maCanHo As String
    Dim ngayKy As Variant
    
    maCanHo = wsData.Range(colCanHo & activeRow).Value
    ngayKy = wsData.Range(colNgayKy & activeRow).Value
    
    '--- Kiem tra dieu kien truoc khi tao so hop dong ---
    If maCanHo <> "" And IsDate(ngayKy) Then
        'Tao chuoi so hop dong theo dung dinh dang
        wsData.Range(colSoHD & activeRow).Value = maCanHo & "/" & Year(ngayKy) & "/2025-HDMB"
    End If
End Sub

