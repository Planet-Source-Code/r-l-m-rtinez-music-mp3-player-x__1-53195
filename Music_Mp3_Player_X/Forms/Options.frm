VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmOpciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Options"
   ClientHeight    =   3375
   ClientLeft      =   2550
   ClientTop       =   4365
   ClientWidth     =   4965
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      Top             =   3030
      UseMaskColor    =   -1  'True
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   12
      Top             =   3030
      Width           =   1305
   End
   Begin VB.FileListBox fileBmps 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   225
      Hidden          =   -1  'True
      Left            =   4965
      Pattern         =   "*.jpg;*.bmp"
      System          =   -1  'True
      TabIndex        =   21
      Top             =   2310
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picContenedor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2475
      Index           =   1
      Left            =   105
      ScaleHeight     =   2475
      ScaleWidth      =   4710
      TabIndex        =   18
      Top             =   480
      Width           =   4710
      Begin VB.OptionButton optWallpaper 
         Caption         =   "No alter."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   465
         TabIndex        =   2
         Top             =   1035
         Width           =   2055
      End
      Begin VB.OptionButton optWallpaper 
         Caption         =   "Tile."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   465
         TabIndex        =   5
         Top             =   2070
         Width           =   1725
      End
      Begin VB.OptionButton optWallpaper 
         Caption         =   "Center."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   465
         TabIndex        =   4
         Top             =   1770
         Width           =   1830
      End
      Begin VB.OptionButton optWallpaper 
         Caption         =   "Strech."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   465
         TabIndex        =   3
         Top             =   1455
         Width           =   2070
      End
      Begin VB.Frame Frame1 
         Caption         =   "Options Wallpaper"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1650
         Left            =   240
         TabIndex        =   19
         Top             =   735
         Width           =   4230
         Begin VB.CheckBox chkProporcional 
            Caption         =   "Proportional"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2115
            TabIndex        =   6
            Top             =   1215
            Width           =   1995
         End
      End
      Begin VB.CheckBox chkDir 
         Alignment       =   1  'Right Justify
         Caption         =   "Enable right click menu in drives and directories"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   105
         Width           =   4140
      End
   End
   Begin VB.PictureBox picContenedor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2475
      Index           =   2
      Left            =   105
      ScaleHeight     =   2475
      ScaleWidth      =   4710
      TabIndex        =   20
      Top             =   480
      Width           =   4710
      Begin VB.ListBox ListaSkins 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2340
         Left            =   90
         TabIndex        =   11
         Top             =   60
         Width           =   4545
      End
   End
   Begin VB.PictureBox picContenedor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2475
      Index           =   4
      Left            =   105
      ScaleHeight     =   2475
      ScaleWidth      =   4710
      TabIndex        =   14
      Top             =   480
      Width           =   4710
      Begin VB.Frame Frame6 
         Caption         =   "Play files"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1155
         Left            =   2085
         TabIndex        =   31
         Top             =   1215
         Width           =   2580
         Begin VB.CheckBox chkWAV 
            Caption         =   " wav files."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   34
            Top             =   780
            Width           =   1980
         End
         Begin VB.CheckBox chkWMA 
            Caption         =   " wma files."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   33
            Top             =   495
            Width           =   1980
         End
         Begin VB.CheckBox chkMP3 
            Caption         =   " mp3 files."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   32
            Top             =   210
            Width           =   1980
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Application"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1155
         Left            =   75
         TabIndex        =   23
         Top             =   60
         Width           =   4590
         Begin VB.CheckBox chkinstancias 
            Caption         =   "Multiple Instances."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   105
            TabIndex        =   10
            Top             =   825
            Width           =   4380
         End
         Begin VB.CheckBox chkSplash 
            Caption         =   "Show Splash Screen."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   105
            TabIndex        =   9
            Top             =   525
            Width           =   4380
         End
         Begin VB.CheckBox chkSiemTop 
            Caption         =   "Always on top."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   105
            TabIndex        =   8
            Top             =   240
            Width           =   4380
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Language"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1155
         Left            =   75
         TabIndex        =   22
         Top             =   1215
         Width           =   1905
         Begin VB.ListBox lstLanguage 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   75
            TabIndex        =   7
            Top             =   210
            Width           =   1755
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mostrar MusicMp3 en"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   840
         Left            =   -75
         TabIndex        =   15
         Top             =   3060
         Visible         =   0   'False
         Width           =   2610
         Begin VB.CheckBox chkMosTask 
            Caption         =   "Barra de Tareas"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   17
            Top             =   240
            Width           =   2310
         End
         Begin VB.CheckBox chkMostTray 
            Caption         =   "System Tray"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   16
            Top             =   495
            Width           =   1590
         End
      End
   End
   Begin VB.PictureBox picContenedor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2475
      Index           =   3
      Left            =   105
      ScaleHeight     =   2475
      ScaleWidth      =   4710
      TabIndex        =   24
      Top             =   480
      Width           =   4710
      Begin VB.Frame Frame3 
         Caption         =   "Alpha (Only win 2000 or later)"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2295
         Left            =   75
         TabIndex        =   25
         Top             =   75
         Width           =   4515
         Begin VB.HScrollBar vScroll1 
            Height          =   255
            Left            =   300
            Max             =   100
            Min             =   10
            TabIndex        =   26
            Top             =   720
            Value           =   100
            Width           =   3885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alpha :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   1095
            TabIndex        =   30
            Top             =   435
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "100%"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   2670
            TabIndex        =   29
            Top             =   435
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100%"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   3885
            TabIndex        =   28
            Top             =   1035
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "10%"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   300
            TabIndex        =   27
            Top             =   1035
            Width           =   315
         End
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   2910
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5133
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Wallpaper"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Skins"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Alpha"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Application"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MiRuta As String

Private Sub chkinstancias_Click()
  If chkinstancias.Value = vbChecked Then
    OpcionesMusic.Instancias = True
  Else
    OpcionesMusic.Instancias = False
  End If
End Sub

Private Sub chkMP3_Click()
 If chkMP3.Value = vbChecked Then
    OpcionesMusic.MP3FILE = True
 Else
    OpcionesMusic.MP3FILE = False
 End If
End Sub

Private Sub chkSiemTop_Click()
  If chkSiemTop.Value = vbChecked Then
    OpcionesMusic.SiempreTop = True
  Else
    OpcionesMusic.SiempreTop = False
  End If
  Always_on_Top
End Sub

Private Sub chkDir_Click()
 On Error Resume Next
  Dim lngRootKey As Long
  Dim RutaExe As String
  lngRootKey = HKEY_CLASSES_ROOT
  
  '+-----------------------------------------------------------------------------------+
  '|procedimento para poner un acceso en el registro para kuando demos click           |
  '|derecho en un folder o driver aparezka el texto 'Buscar Music Mp3' y se ejecute la |
  '|aplicacion con los parametros enviados en este caso donde dimos click derecho      |
  '|las claves son:                                                                    |
  '| --> HKEY_CLASSES_ROOT\Directory\Shell\ 'Texto del Menu'                           |
  '| --> HKEY_CLASSES_ROOT\Directory\Shell\ 'Texto del Menu' \command                  |
  '|                                  con una clave con la ruta de la aplicacion y     |
  '|                                  comandos                                         |
  '+-----------------------------------------------------------------------------------+
  
  If chkDir.Value = vbChecked Then
    OpcionesMusic.Directorio = True
    '// obtener la string correcta para ponerla en el registro
    RutaExe = Path_Exe(PathExe) & App.EXEName & ".exe %1"
     'Verifikar si existe la clave
    If Not regDoes_Key_Exist(lngRootKey, "Directory\shell\Search Music Mp3 Player X") Then
      regCreate_A_Key lngRootKey, "Directory\shell\Search Music Mp3 Player X"
      regCreate_A_Key lngRootKey, "Directory\shell\Search Music Mp3 Player X\command"
      regCreate_Key_Value lngRootKey, "Directory\shell\Search Music Mp3 Player X\command", "", RutaExe
    End If
    If Not regDoes_Key_Exist(lngRootKey, "Drive\shell\Search Music Mp3 Player X") Then
      regCreate_A_Key lngRootKey, "Drive\shell\Search Music Mp3 Player X"
      regCreate_A_Key lngRootKey, "Drive\shell\Search Music Mp3 Player X\command"
      regCreate_Key_Value lngRootKey, "Drive\shell\Search Music Mp3 Player X\command", "", RutaExe
    End If
  Else
     OpcionesMusic.Directorio = False
     regDelete_A_Key lngRootKey, "Directory\shell\Search Music Mp3 Player X", "command"
     regDelete_A_Key lngRootKey, "Directory\shell", "Search Music Mp3 Player X"
     regDelete_A_Key lngRootKey, "Drive\shell\Search Music Mp3 Player X", "command"
     regDelete_A_Key lngRootKey, "Drive\shell", "Search Music Mp3 Player X"
  End If
End Sub

Private Sub chkMosTask_Click()
'  If chkMosTask.Value = vbChecked Then
'    MusicMp3.ShowInTaskbar = True
'  Else
'    MusicMp3.ShowInTaskbar = False
'  End If
End Sub

Private Sub chkProporcional_Click()
  If chkProporcional.Value = vbChecked Then
    OpcionesMusic.Proporcional = True
  Else
    OpcionesMusic.Proporcional = False
  End If
End Sub

Private Sub chkSplash_Click()
  If chkSplash.Value = vbChecked Then
    OpcionesMusic.Splash = True
  Else
    OpcionesMusic.Splash = False
  End If
End Sub


Private Sub Apply_Skin()
 On Error Resume Next
 Dim Skins As String, aryName() As String
 Dim i As Integer

If ListaSkins.ListIndex <= 0 Then Exit Sub

cmdApply.Enabled = False
Skins = Trim(ListaSkins.Text)
Skins = Right(Skins, Len(Skins) - 4)
aryName = Split(Skins, "\", , vbTextCompare)

'// obtener el nombre del skin
If UBound(aryName) <> 0 Then
   Skins = Trim(aryName(0))
Else
   Skins = Trim(aryName(0))
End If

'// si es el mismo skin no kambiarlo
If LCase(Skins) = LCase(strSkinSeleccionado) Then: cmdApply.Enabled = True: Exit Sub


If Skins = "Default" Then
   strSkinSeleccionado = "\" & Skins
    '// si esta la minimascara
    If bolMiniMascara = True Then
       frmMini.Visible = False
    Else
       MusicMp3.Visible = False
    End If
 
    '// seleccionar el menu correcto del skin
    For i = 1 To frmPopUp.mnuSkinsAdd.Count
       If LCase(Trim(frmPopUp.mnuSkinsAdd(i).Caption)) = strSkinSeleccionado Then
          frmPopUp.mnuSkinsAdd(i).Checked = True
       Else
          frmPopUp.mnuSkinsAdd(i).Checked = False
       End If
    Next i
  
    '// cambiar el skin
    Change_Skin strSkinSeleccionado
    '// ajustar los bordes
    Form_Mini_Normal
    '// si esta la minimascara
    If bolMiniMascara = True Then
    frmMini.Visible = True
      'Change_Mask True
    Else
      MusicMp3.Visible = True
      'Change_Mask False
    End If
    cmdApply.Enabled = True
    Exit Sub
End If

'// chekar si existe la carpeta
If Dir(MiRuta & Skins, vbDirectory) <> "" Then
    strSkinSeleccionado = Skins
    '// si esta la minimascara
    If bolMiniMascara = True Then
       frmMini.Visible = False
    Else
       MusicMp3.Visible = False
    End If

    '// seleccionar el menu correcto del skin
    For i = 1 To frmPopUp.mnuSkinsAdd.Count
      If LCase(Trim(frmPopUp.mnuSkinsAdd(i).Caption)) = LCase(strSkinSeleccionado) Then
         frmPopUp.mnuSkinsAdd(i).Checked = True
      Else
         frmPopUp.mnuSkinsAdd(i).Checked = False
      End If
    Next i
    
    '// Cambiar el skin
    Change_Skin Skins
    '// ajustar los bordes
    Form_Mini_Normal
    
    If bolMiniMascara = True Then
      'Change_Mask True
      frmMini.Visible = True
    Else
      MusicMp3.Visible = True
      'Change_Mask False
    End If
End If

cmdApply.Enabled = True
End Sub



Private Sub chkWAV_Click()
 If chkWAV.Value = vbChecked Then
    OpcionesMusic.WAVFILE = True
 Else
    OpcionesMusic.WAVFILE = False
 End If

End Sub

Private Sub chkWMA_Click()
 If chkWMA.Value = vbChecked Then
    OpcionesMusic.WMAFILE = True
 Else
    OpcionesMusic.WMAFILE = False
 End If

End Sub

Private Sub cmdApply_Click()
 Select Case TabStrip1.SelectedItem.Index
   Case 1
      If frmPopUp.mnuWallpapper = True Then ConfigurarWallpaper
   Case 2
      Apply_Skin
   Case 4
      Load_Language OpcionesMusic.Language
     If chkMP3.Value = vbUnchecked And chkWMA.Value = vbUnchecked And chkWAV.Value = vbUnchecked Then
       chkMP3.Value = vbChecked
       OpcionesMusic.MP3FILE = True
     End If
 End Select

End Sub

Private Sub cmdCancel_Click()
   If TabStrip1.SelectedItem.Index = 4 Then
     If chkMP3.Value = vbUnchecked And chkWMA.Value = vbUnchecked And chkWAV.Value = vbUnchecked Then
       OpcionesMusic.MP3FILE = True
     End If
   End If
  Unload Me
End Sub

Private Sub Form_Load()
 Dim miNombre As String, strInfo As String, strSkinTemp As String
 Dim i As Integer
  On Error Resume Next
  bolOpcionesShow = True
  Load_Language_Options '// cargar lenguaje siempre

  MiRuta = Path_Exe(PathExe) & "MMp3Player\Skins\"
 'configuration options wallpaper
 optWallpaper(0).Value = OpcionesMusic.NoAlteraR: optWallpaper(1).Value = OpcionesMusic.Mosaico
 optWallpaper(2).Value = OpcionesMusic.Centrar: optWallpaper(3).Value = OpcionesMusic.Expander
 'center form
 Me.left = (Screen.Width - Me.Width) / 2: Me.Top = (Screen.Height - Me.Height) / 2
 'configuration checks
 If OpcionesMusic.Proporcional = True Then chkProporcional.Value = vbChecked
 If OpcionesMusic.Splash = True Then chkSplash.Value = vbChecked
 If OpcionesMusic.Instancias = True Then chkinstancias.Value = vbChecked
 If OpcionesMusic.Directorio = True Then chkDir.Value = vbChecked
 If OpcionesMusic.SiempreTop = True Then chkSiemTop.Value = vbChecked
 
 If OpcionesMusic.MP3FILE = True Then chkMP3.Value = vbChecked
 If OpcionesMusic.WMAFILE = True Then chkWMA.Value = vbChecked
 If OpcionesMusic.WAVFILE = True Then chkWAV.Value = vbChecked
 
 vScroll1.Value = OpcionesMusic.Alpha
 Label1(3).Caption = OpcionesMusic.Alpha & "%"
'-------------------------------------------------------------------------
     ListaSkins.BackColor = Read_INI("Skin", "RepBackColor", RGB(0, 0, 0), True)
     ListaSkins.ForeColor = Read_INI("Skin", "RepForeColor", RGB(255, 255, 255), True)
     lstLanguage.BackColor = Read_INI("Skin", "RepBackColor", RGB(0, 0, 0), True)
     lstLanguage.ForeColor = Read_INI("Skin", "RepForeColor", RGB(255, 255, 255), True)
'-------------------------------------------------------------------------

'search skins in musicmp3/skins only directories
 ListaSkins.Clear
 ListaSkins.AddItem "+-Skins  "
 ListaSkins.AddItem " +-> Default \ By R@úL M@RtInEz" '// Agregar siempre el de deafult
 ListaSkins.Selected(1) = True      '// y seleccionarlo
 miNombre = Dir(MiRuta, vbDirectory) '// recuperar la primera entrada en la ruta
 i = 1
 strSkinTemp = strSkinSeleccionado
 Do While miNombre <> ""
   If miNombre <> "." And miNombre <> ".." Then
      ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
      If (GetAttr(MiRuta & miNombre) And vbDirectory) = vbDirectory Then
        fileBmps.Path = MiRuta & miNombre
       '// chekar si hay archivos jpg o bmps pára ponerlos como posible skin
       If fileBmps.ListCount > 0 Then
        strSkinSeleccionado = miNombre
        strInfo = Read_INI("Info", "AuthorName", "")  '// Obtener el autor del skin
         i = i + 1
          If strInfo = "" Then
             ListaSkins.AddItem " +-> " & miNombre
          Else
            ListaSkins.AddItem " +-> " & miNombre & " \ By " & strInfo
          End If
          
          '// Seleccionar el skin actual si esta
          If LCase(Trim(miNombre)) = LCase(Trim(strSkinTemp)) Then ListaSkins.Selected(i) = True
        End If
      End If
   End If
   miNombre = Dir
 Loop
strSkinSeleccionado = strSkinTemp

'-----------------------------------------------------------------------------------
'// buskar los archivos de lenguaje y agragarlos
miNombre = Dir(Path_Exe(PathExe) & "MMp3Player\Language\*.*")    ' Recupera la primera entrada.
i = 0
Do While miNombre <> ""
   If miNombre <> "." And miNombre <> ".." Then
      ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
        If Right(LCase(miNombre), 3) = "lng" Then  '// verifikar la extencion del archivo
           strInfo = left(Trim(miNombre), Len(Trim(miNombre)) - 4)
           lstLanguage.AddItem strInfo
         '// Seleccionar el lenguaje que se esta utilizando
         If LCase(Trim(strInfo)) = LCase(Trim(OpcionesMusic.Language)) Then
            lstLanguage.Selected(i) = True
         End If
         i = i + 1
        End If
   End If
   miNombre = Dir
Loop
 
 '// si no hay ningun archivo de lenguage poner el de defaul Español :P
 If lstLanguage.ListCount = 0 Then
   lstLanguage.AddItem "English"
   lstLanguage.Selected(0) = True
 End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
  bolOpcionesShow = False
End Sub

Private Sub lstLanguage_Click()
 If lstLanguage.ListCount = 0 Then Exit Sub
  OpcionesMusic.Language = Trim(lstLanguage.List(lstLanguage.ListIndex))
End Sub


Private Sub optWallpaper_Click(Index As Integer)
  
  If optWallpaper(0).Value = True Or optWallpaper(3).Value = True Then
    chkProporcional.Value = vbUnchecked
    chkProporcional.Enabled = False
    OpcionesMusic.NoAlteraR = optWallpaper(0).Value
    OpcionesMusic.Expander = optWallpaper(3).Value
    OpcionesMusic.Mosaico = False
    OpcionesMusic.Centrar = False
  Else
    OpcionesMusic.Mosaico = optWallpaper(1).Value
    OpcionesMusic.Centrar = optWallpaper(2).Value
    OpcionesMusic.Expander = False
    OpcionesMusic.NoAlteraR = False
    chkProporcional.Enabled = True
  End If
End Sub

Private Sub TabStrip1_Click()
  picContenedor(TabStrip1.SelectedItem.Index).ZOrder vbBringToFront
End Sub

Private Sub VScroll1_Scroll()
 On Error GoTo Hell
   Dim i As Integer
   '// Ajustar a porcentaje
   Label1(3).Caption = (vScroll1.Value * 100) / 100 & "%"
       If bolMiniMascara = True Then
        HaceR_TransparentE frmMini.hwnd, vScroll1.Value
       Else
        HaceR_TransparentE MusicMp3.hwnd, vScroll1.Value
       End If
        OpcionesMusic.Alpha = vScroll1.Value
      
      For i = 0 To 9 '// deseleccionar los menus de porcentaje
        frmPopUp.mnuAlpha(i).Checked = False
      Next i
        '// seleccionar el menu de personalizado y  poner porcentaje
        frmPopUp.mnuAlphaPer.Caption = Trim(arryLanguage(30)) & " [ " & vScroll1.Value & "% ]"
        frmPopUp.mnuAlphaPer.Checked = True
 Exit Sub
Hell:
End Sub
