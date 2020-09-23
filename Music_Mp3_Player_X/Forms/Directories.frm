VERSION 5.00
Begin VB.Form frmDirectorios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Albums Browser"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "Directories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstAlbums 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2340
      IntegralHeight  =   0   'False
      ItemData        =   "Directories.frx":000C
      Left            =   0
      List            =   "Directories.frx":000E
      TabIndex        =   0
      Top             =   0
      Width           =   4710
   End
End
Attribute VB_Name = "frmDirectorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Load()
On Error Resume Next
 Dim ruta As String
 Dim Discos As Integer, i As Integer
 Dim a As String
  bolDirectoriosShow = True
  Me.Caption = arryLanguage(6) & " [ " & TotalAlbumS & " Albums ]"
  
  frmDirectorios.left = (Screen.Width - frmDirectorios.Width) / 2
  frmDirectorios.Top = (Screen.Height - frmDirectorios.Height) / 2

'---------------------------------------------------------------------------------------
     lstAlbums.BackColor = Read_INI("Skin", "RepBackColor", RGB(0, 0, 0), True)
     lstAlbums.ForeColor = Read_INI("Skin", "RepForeColor", RGB(255, 255, 255), True)
'---------------------------------------------------------------------------------------
 Load_Albums
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Sub Load_Albums()
 Dim ruta As String, a As String
 Dim Discos As Integer, i As Integer, j As Integer
 Dim Arreglo() As String

On Error Resume Next
 
 Discos = TotalAlbumS  '// almacenar los albums totales buskados
 'ruta = MusicMp3.picAlbum(1).ToolTipText  '// almacenar la ruta del primer album
 ruta = strTraySearch
 lstAlbums.Clear
 'lstAlbums.AddItem ">> " & ruta
 For i = 1 To Discos
   a = Mid(MusicMp3.picAlbum(i).ToolTipText, Len(ruta), Len(MusicMp3.picAlbum(i).ToolTipText))
    If Trim(a) <> "" Then
      lstAlbums.AddItem a
    Else
      Arreglo = Split(MusicMp3.picAlbum(i).ToolTipText, "\")
      j = UBound(Arreglo)
      lstAlbums.AddItem "\" & Trim(Arreglo(j))
    End If
 Next i
  '// seleccionar el album reproduciendo
  lstAlbums.Selected(intActiveAlbum - 1) = True
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Resize()
  lstAlbums.Width = Me.ScaleWidth
  lstAlbums.Height = Me.ScaleHeight
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Unload(Cancel As Integer)
bolDirectoriosShow = False
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub lstAlbums_DblClick()
If lstAlbums.ListCount = 0 Then Exit Sub
 MusicMp3.ListaRep.ListIndex = -1
 MusicMp3.Album_Reproducir lstAlbums.ListIndex + 1

End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
