VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Awal Kalimat Kapital"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function AwalKalimatKapital(strKalimat As String)
Dim Temp1 As String, Temp2 As String
Dim Lokasi As Integer, i As Integer
Dim huruf As String * 1
  Temp1$ = LCase(strKalimat)  'Kecilkan dulu semua
  For i% = 1 To Len(Temp1$)
    huruf = Chr(Asc(Mid(strKalimat, i%, 1)))
    If huruf = "." Then
      Lokasi% = i% + 2
    End If
    If i% = 1 Or i% = Lokasi% Then
       Temp2$ = Temp2$ + UCase(Chr(Asc(Mid(Temp1$, i%, 1))))
    Else
       Temp2$ = Temp2$ + LCase(Chr(Asc(Mid(Temp1$, i%, 1))))
    End If
  Next i
  AwalKalimatKapital = Temp2$
End Function

Private Sub Text1_Change()
  Dim posisi As Integer
  posisi = Text1.SelStart
  Text1.Text = AwalKalimatKapital(Text1.Text)
  Text1.SelStart = posisi
End Sub


