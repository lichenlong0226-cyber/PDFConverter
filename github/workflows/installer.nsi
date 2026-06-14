; installer.nsi
OutFile "release_artifacts\\PDFConverter-setup.exe"
InstallDir "$PROGRAMFILES\\PDFConverter"

Page directory
Page instfiles

Section "Install"
  SetOutPath "$INSTDIR"
  File "release_artifacts\\PDFConverter.exe"
  CreateShortCut "$DESKTOP\\PDFConverter.lnk" "$INSTDIR\\PDFConverter.exe"
SectionEnd

Section "Uninstall"
  Delete "$INSTDIR\\PDFConverter.exe"
  Delete "$DESKTOP\\PDFConverter.lnk"
  RMDir "$INSTDIR"
SectionEnd
