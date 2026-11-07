; Grimoire — compléments de l'installateur Windows (inclus par electron-builder).
; Déclare Grimoire dans Paramètres > Applications > Applications par défaut, comme
; LibreOffice : un clic sur « Définir par défaut » et les .docx, .xlsx, .csv et .pdf
; s'ouvrent avec Grimoire, chacun avec son icône (DOC, XLS, CSV, PDF).

!macro customInstall
  ; Restes de l'ancien script de la version navigateur (grimoire-icones / grimoire-installation.cmd)
  DeleteRegKey HKCU "Software\Classes\Grimoire.Document\Application"
  DeleteRegKey HKCU "Software\Classes\Grimoire.Classeur\Application"
  DeleteRegKey HKCU "Software\Classes\Grimoire.CSV\Application"
  DeleteRegKey HKCU "Software\Classes\Grimoire.PDF\Application"
  DeleteRegKey HKCU "Software\Grimoire-icones"

  WriteRegStr HKCU "Software\Grimoire\Capabilities" "ApplicationName" "Grimoire"
  WriteRegStr HKCU "Software\Grimoire\Capabilities" "ApplicationDescription" "Suite bureautique personnelle : documents, classeurs, PDF."
  WriteRegStr HKCU "Software\Grimoire\Capabilities" "ApplicationIcon" "$appExe,0"
  WriteRegStr HKCU "Software\Grimoire\Capabilities\FileAssociations" ".docx" "Grimoire.Document"
  WriteRegStr HKCU "Software\Grimoire\Capabilities\FileAssociations" ".xlsx" "Grimoire.Classeur"
  WriteRegStr HKCU "Software\Grimoire\Capabilities\FileAssociations" ".xlsm" "Grimoire.Classeur"
  WriteRegStr HKCU "Software\Grimoire\Capabilities\FileAssociations" ".csv" "Grimoire.CSV"
  WriteRegStr HKCU "Software\Grimoire\Capabilities\FileAssociations" ".pdf" "Grimoire.PDF"
  WriteRegStr HKCU "Software\RegisteredApplications" "Grimoire" "Software\Grimoire\Capabilities"
  ; nom affiché dans « Ouvrir avec »
  WriteRegStr HKCU "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}" "FriendlyAppName" "Grimoire"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro customUnInstall
  DeleteRegValue HKCU "Software\RegisteredApplications" "Grimoire"
  DeleteRegKey HKCU "Software\Grimoire"
  DeleteRegKey HKCU "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
