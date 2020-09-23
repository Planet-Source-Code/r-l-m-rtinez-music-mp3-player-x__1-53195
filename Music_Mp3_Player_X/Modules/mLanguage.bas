Attribute VB_Name = "mLanguage"
Option Explicit

Public arryLanguage() As String

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  IDIOMA                                                                               |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


Public Sub Load_Language(strLang As String)
 On Error Resume Next
 Dim Linenr
 Dim InputData
 Dim i As Integer
 Dim strRuta As String, strTemp As String
 ReDim arryLanguage(65)
 arryLanguage(1) = "    New Find"
 arryLanguage(2) = "    Change ListRep/Cover Front"
 arryLanguage(3) = "    Put Cover Front as Wallpaper"
 arryLanguage(4) = "    Maximize Cover front"
 arryLanguage(5) = "    Explore Mp3's"
 arryLanguage(6) = "    Albums Browser"
 arryLanguage(7) = "    Players Controls"
 arryLanguage(8) = "  Volume"
 arryLanguage(9) = "+   Increase Volume"
 arryLanguage(10) = "-   Decrease Volume"
 arryLanguage(11) = "Z   Previous Track"
 arryLanguage(12) = "X   Play"
 arryLanguage(13) = "C   Pause"
 arryLanguage(14) = "V   Stop"
 arryLanguage(15) = "B   Next Track"
 arryLanguage(16) = "<   Previous Album/Folder"
 arryLanguage(17) = ">   Next Album/Folder"
 arryLanguage(18) = "I   Intro 10 seg."
 arryLanguage(19) = "R   Repeat Track"
 arryLanguage(20) = "S   Mute"
 arryLanguage(21) = "  Randomize"
 arryLanguage(22) = "Q   Current Album/Folder"
 arryLanguage(23) = "W   Whole Albums"
 arryLanguage(24) = "A   Skip Backward 5 sec."
 arryLanguage(25) = "D   Skip Forward 5 sec."
 arryLanguage(26) = "    Options..."
 arryLanguage(27) = "    Skins..."
 arryLanguage(28) = " << Skins Browser >>"
 arryLanguage(29) = "    Alpha Mode"
 arryLanguage(30) = " Custom"
 arryLanguage(31) = "    About..."
 arryLanguage(32) = "    Minimize"
 arryLanguage(33) = "    Change Mask"
 arryLanguage(34) = "    Exit"
 arryLanguage(35) = " Wallpaper"
 arryLanguage(36) = " Skins"
 arryLanguage(37) = " Alpha"
 arryLanguage(38) = " Application"
 arryLanguage(39) = " Enable right click menu in drives and directories"
 arryLanguage(40) = " Options Wallpaper"
 arryLanguage(41) = " No Alter."
 arryLanguage(42) = " Strech."
 arryLanguage(43) = " Center."
 arryLanguage(44) = " Tile."
 arryLanguage(45) = " Proportional."
 arryLanguage(46) = " Alpha (Only win 2000 or later.)"
 arryLanguage(47) = " Alpha: "
 arryLanguage(48) = " Language"
 arryLanguage(49) = " Application"
 arryLanguage(50) = " Always on Top."
 arryLanguage(51) = " Show Splash Screen."
 arryLanguage(52) = " Multiple Instances."
 arryLanguage(53) = " Play Files"
 arryLanguage(54) = " .mp3 Files."
 arryLanguage(55) = " .wma Files."
 arryLanguage(56) = " .wav Files."
 arryLanguage(57) = " [ Searching... ]"
 arryLanguage(58) = " [ No Mp3's files found ]"
 arryLanguage(59) = " Apply"
 arryLanguage(60) = " Cancel"
 arryLanguage(61) = " error reading file"
 arryLanguage(62) = " Current Cover Front"
 arryLanguage(63) = " Searching files in:"
 arryLanguage(64) = " Select a directory for search."

   
  strRuta = Path_Exe(PathExe) & "MMp3Player\Language\" & strLang & ".lng"
   If Dir(strRuta) <> "" Then
    Open strRuta For Input As #2

     Linenr = -1
     Do While Not EOF(2)
       Line Input #2, InputData
        i = i + 1
        If i > 65 Then Exit Do
        If Trim(InputData) <> "" Or Len(Trim(InputData)) > 3 Then
          Linenr = Linenr + 1
          strTemp = left(arryLanguage(Linenr), 1)
          strTemp = Trim(strTemp) & "   " & InputData
          arryLanguage(Linenr) = strTemp
          If Linenr > 34 Then arryLanguage(Linenr) = Trim(strTemp)
        End If
     Loop
    Close #2
   End If
 With frmPopUp
   .mnuNuevaBusqueda.Caption = arryLanguage(1)
   .mnuCambiarListaCaratula.Caption = arryLanguage(2)
   MusicMp3.imgNormal(10).ToolTipText = Trim(arryLanguage(2))
   .mnuWallpapper.Caption = arryLanguage(3)
   .mnuMCaratula.Caption = arryLanguage(4)
   .mnuExplorar.Caption = arryLanguage(5)
   .mnuExpAlbum.Caption = arryLanguage(6)
   .mnuControles.Caption = arryLanguage(7)
   .mnuVolumen.Caption = arryLanguage(8)
   .mnuSubirVolumen.Caption = arryLanguage(9)
   .mnuBajarVolumen.Caption = " " & arryLanguage(10)
   .mnuTrackAnterior.Caption = arryLanguage(11)
   MusicMp3.imgNormal(0).ToolTipText = Trim(Right(arryLanguage(11), Len(arryLanguage(11)) - 1))
   frmMini.picNormal(0).ToolTipText = Trim(Right(arryLanguage(11), Len(arryLanguage(11)) - 1))
   .mnuReproducir.Caption = arryLanguage(12)
   MusicMp3.imgNormal(1).ToolTipText = Trim(Right(arryLanguage(12), Len(arryLanguage(12)) - 1))
   frmMini.picNormal(1).ToolTipText = Trim(Right(arryLanguage(12), Len(arryLanguage(12)) - 1))
   .mnuPausa.Caption = arryLanguage(13)
   MusicMp3.imgNormal(2).ToolTipText = Trim(Right(arryLanguage(13), Len(arryLanguage(13)) - 1))
   frmMini.picNormal(2).ToolTipText = Trim(Right(arryLanguage(13), Len(arryLanguage(13)) - 1))
   .mnuDetener.Caption = arryLanguage(14)
   MusicMp3.imgNormal(3).ToolTipText = Trim(Right(arryLanguage(14), Len(arryLanguage(14)) - 1))
   frmMini.picNormal(3).ToolTipText = Trim(Right(arryLanguage(14), Len(arryLanguage(14)) - 1))
   .mnuSigTrack.Caption = arryLanguage(15)
   MusicMp3.imgNormal(4).ToolTipText = Trim(Right(arryLanguage(15), Len(arryLanguage(15)) - 1))
   frmMini.picNormal(4).ToolTipText = Trim(Right(arryLanguage(15), Len(arryLanguage(15)) - 1))
   .mnuAnteriorAlbum.Caption = arryLanguage(16)
   MusicMp3.imgNormal(9).ToolTipText = Trim(Right(arryLanguage(16), Len(arryLanguage(16)) - 1))
   .mnuSigAlbum.Caption = arryLanguage(17)
   MusicMp3.imgNormal(11).ToolTipText = Trim(Right(arryLanguage(17), Len(arryLanguage(17)) - 1))
   .mnuIntro.Caption = arryLanguage(18)
   MusicMp3.imgNormal(5).ToolTipText = Trim(Right(arryLanguage(18), Len(arryLanguage(18)) - 1))
   .mnuSilencio.Caption = arryLanguage(20)
   MusicMp3.imgNormal(6).ToolTipText = Trim(Right(arryLanguage(20), Len(arryLanguage(20)) - 1))
   .mnuRepetir.Caption = arryLanguage(19)
   MusicMp3.imgNormal(7).ToolTipText = Trim(Right(arryLanguage(19), Len(arryLanguage(19)) - 1))
   .mnuOrdenAleatorio.Caption = arryLanguage(21)
   MusicMp3.imgNormal(8).ToolTipText = Trim(arryLanguage(21))
   .mnuAleatorioActAlbum.Caption = arryLanguage(22)
   .mnuAleatorioTodaColec.Caption = arryLanguage(23)
   .mnuAtras5Seg.Caption = arryLanguage(24)
   .mnuAdelante5Seg.Caption = arryLanguage(25)
   .mnuOpciones.Caption = arryLanguage(26)
   .mnuSkins.Caption = arryLanguage(27)
   .mnuExpSkins.Caption = arryLanguage(28)
   .mnuWOpacity.Caption = arryLanguage(29)
   .mnuAlphaPer.Caption = Trim(arryLanguage(30))
   .mnuAcercaDe.Caption = arryLanguage(31)
   .mnuMinimizar.Caption = arryLanguage(32)
   .mnuCambiarMascaras.Caption = arryLanguage(33)
   .mnuSalir.Caption = arryLanguage(34)
   MusicMp3.imgNormal(12).ToolTipText = Trim(arryLanguage(32))
   MusicMp3.imgNormal(13).ToolTipText = Trim(arryLanguage(33))
   frmMini.picNormal(5).ToolTipText = Trim(arryLanguage(33))
   MusicMp3.imgNormal(14).ToolTipText = Trim(arryLanguage(34))
   frmMini.picNormal(6).ToolTipText = Trim(arryLanguage(34))
   If bolOpcionesShow = True Then Load_Language_Options
   If bolDirectoriosShow = True Then frmDirectorios.Caption = arryLanguage(6) & " [ " & TotalAlbumS & " Albums ]"
   If bolCaratulaShow = True Then frmCaratula.Caption = arryLanguage(62)
   If bolAcercaShow = True Then frmAcerca.Caption = arryLanguage(31)
 End With
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Load_Language_Options()
 With frmOpciones
   .Caption = arryLanguage(26)
   .TabStrip1.Tabs(1).Caption = arryLanguage(35)
   .TabStrip1.Tabs(2).Caption = arryLanguage(36)
   .TabStrip1.Tabs(3).Caption = arryLanguage(37)
   .TabStrip1.Tabs(4).Caption = arryLanguage(38)
   .chkDir.Caption = arryLanguage(39)
   '//walpaper
   .Frame1.Caption = arryLanguage(40)
   .optWallpaper(0).Caption = arryLanguage(41)
   .optWallpaper(3).Caption = arryLanguage(42)
   .optWallpaper(2).Caption = arryLanguage(43)
   .optWallpaper(1).Caption = arryLanguage(44)
   .chkProporcional.Caption = arryLanguage(45)
   '//alpha
   .Frame3.Caption = arryLanguage(46)
   .Label1(2).Caption = arryLanguage(47)
   '//language
   .Frame2.Caption = arryLanguage(48)
   '// application
   .Frame5.Caption = arryLanguage(49)
   .chkSiemTop.Caption = arryLanguage(50)
   .chkSplash.Caption = arryLanguage(51)
   .chkinstancias.Caption = arryLanguage(52)
   '// format files
   .Frame6.Caption = arryLanguage(53)
   .chkMP3.Caption = arryLanguage(54)
   .chkWMA.Caption = arryLanguage(55)
   .chkWAV.Caption = arryLanguage(56)
   
   '//buttons
   .cmdApply.Caption = arryLanguage(59)
   .cmdCancel.Caption = arryLanguage(60)
  End With
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

