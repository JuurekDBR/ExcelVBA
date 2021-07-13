VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Punkty"
   ClientHeight    =   1836
   ClientLeft      =   96
   ClientTop       =   396
   ClientWidth     =   2184
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" _
        (ByVal nIndex As LongPtr) As LongPtr
Private Sub CommandButton1_Click()
    poziomo = GetSystemMetrics(16)
    pasek = GetSystemMetrics(4)
    pionowo = GetSystemMetrics(17) + pasek
    With ActiveWindow.ActivePane
        punktT = .PointsToScreenPixelsX(3)
        punkt0 = .PointsToScreenPixelsX(0)
    End With
    dpi = (punktT - punkt0) / 3 * 72     'nie zmienia w czasie pracy arkusza
    czynnik = 72 / dpi                   'mój mno¿nik do zamiany
    Me.Top = 0: Me.Left = 0
    Me.Height = pionowo * czynnik
    Me.Width = poziomo * czynnik
    Label1.Caption = "Punkty EXCEL: " & Me.Width & " x " & Me.Height & " (" & dpi & ") " & czynnik
End Sub
Private Sub Label1_Click()
End Sub
Private Sub UserForm_Click()
End Sub

