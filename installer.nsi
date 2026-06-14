; installer.nsi
OutFile "release_artifacts\\PDFConverter-setup.exe"
InstallDir "$LOCALAPPDATA\\PDFConverter"
RequestExecutionLevel admin

Page directory
Page instfiles

Section "Install"
  SetOutPath "$INSTDIR"
  File "release_artifacts\\PDFConverter.exe"
  CreateShortCut "$DESKTOP\\PDFConverter.lnk" "$INSTDIR\\PDFConverter.exe"
  WriteUninstaller "$INSTDIR\\uninstall.exe"
SectionEnd

Section "Uninstall"
  Delete "$INSTDIR\\PDFConverter.exe"
  Delete "$INSTDIR\\uninstall.exe"
  Delete "$DESKTOP\\PDFConverter.lnk"
  RMDir "$INSTDIR"
SectionEnd

