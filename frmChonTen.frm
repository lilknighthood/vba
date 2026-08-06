VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChonTen 
   Caption         =   "Chon ten"
   ClientHeight    =   756
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6384
   OleObjectBlob   =   "frmChonTen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChonTen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Khai bao API de ep Title Form hien thi Unicode bang lenh he thong loi cua Windows
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function DefWindowProcW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function DefWindowProcW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

' Ma lenh yeu cau Windows thay doi Tieu de cua so
Private Const WM_SETTEXT As Long = &HC

Public TenDuocChon As String
Public DaHuy As Boolean

Private Sub UserForm_Initialize()
    Dim wsSetup As Worksheet
    Dim iRow As Long
    
    ' --- 1. SET TI?NG VI?T CHO NÚT B?M (Dùng ChrW) ---
    btnOK.Caption = "Ch" & ChrW(7885) & "n"
    btnCancel.Caption = "H" & ChrW(7911) & "y"
    
    ' Ð?t m?t cái tên t?m không d?u d? hàm API tìm th?y ID c?a s?
    Me.Caption = "FormXuatFile"
    
    ' --- 2. Ð?C D? LI?U T? SETUP ---
    Set wsSetup = ThisWorkbook.Sheets("Setup")
    iRow = 6
    cmbTen.Clear
    
    While Trim(wsSetup.Range("M" & iRow).Value) <> ""
        cmbTen.AddItem Trim(wsSetup.Range("M" & iRow).Value)
        iRow = iRow + 1
    Wend
    
    If cmbTen.ListCount > 0 Then cmbTen.ListIndex = 0
    DaHuy = False
End Sub

' --- S? D?NG S? KI?N ACTIVATE & DEFWINDOWPROCW ---
Private Sub UserForm_Activate()
    #If VBA7 Then
        Dim hwnd As LongPtr
    #Else
        Dim hwnd As Long
    #End If
    
    ' Tìm ID c?a Form d?a vào tên t?m
    hwnd = FindWindow("ThunderDFrame", "FormXuatFile")
    
    ' Dùng ChrW ghép chu?i d? trình so?n th?o VBA không làm h?ng font
    Dim titleUni As String
    titleUni = "Ch" & ChrW(7885) & "n t" & ChrW(234) & "n file mu" & ChrW(7889) & "n xu" & ChrW(7845) & "t"
    
    ' Ép Windows hi?n th? chu?i Unicode b?ng l?nh lõi (WM_SETTEXT)
    If hwnd <> 0 Then
        DefWindowProcW hwnd, WM_SETTEXT, 0, StrPtr(titleUni)
    End If
End Sub

Private Sub btnOK_Click()
    TenDuocChon = cmbTen.Value
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    DaHuy = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        DaHuy = True
        Me.Hide
    End If
End Sub
