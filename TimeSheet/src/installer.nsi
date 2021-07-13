!include "FileFunc.nsh"
!include "WordFunc.nsh"
!include "LogicLib.nsh"
;!include "version.txt"

;NSIS Modern User Interface
;Start Menu Folder Selection Example Script

!define version "1.0"
VIProductVersion "1.0.0.0"

;--------------------------------
;Include Modern UI

  !include "MUI2.nsh"
 
;--------------------------------
;General

!define MUI_ICON "src\TSA.ico"
!define MUI_UNICON "src\TSA.ico"
!define SHORT_NAME "TSA"
!define FULL_NAME "TimeSheet Application"

  ;Name and file
  Name "TSA"
;  OutFile "${SHORT_NAME} v${version}.exe"
  OutFile "${SHORT_NAME}.exe"

  ;Default installation folder
  InstallDir "$PROGRAMFILES\${SHORT_NAME}"
  
  ;Get installation folder from registry if available
  InstallDirRegKey HKCU "Software\${SHORT_NAME}\" "Install Location"

  ;Request application privileges for Windows Vista
  RequestExecutionLevel user

;    VIProductVersion "${version}"
    VIAddVersionKey ProductName "${SHORT_NAME}"
    VIAddVersionKey CompanyName "Northrop Grumman Corporation"
    VIAddVersionKey FileDescription "${SHORT_NAME} is a tool to analyze CATS charges"
    VIAddVersionKey LegalCopyright "© 2012 Northrop Grumman Corporation "
    VIAddVersionKey FileVersion "${version}"
    VIAddVersionKey ProductVersion "${version}"

    CRCCheck force
    BrandingText "${FULL_NAME} v${version}"
;    BrandingText "${FULL_NAME}"
  
;--------------------------------
;Variables
  Var RegistryVersion
  Var StartMenuFolder
  
;--------------------------------
;Interface Settings

  !define MUI_ABORTWARNING

;--------------------------------
;Pages

  #!insertmacro MUI_PAGE_LICENSE "${NSISDIR}\Docs\Modern UI\License.txt"
  #!insertmacro MUI_PAGE_COMPONENTS
  !insertmacro MUI_PAGE_DIRECTORY
  
  ;Start Menu Folder Page Configuration
  !define MUI_STARTMENUPAGE_REGISTRY_ROOT "HKCU" 
  !define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\${SHORT_NAME}" 
  !define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "${SHORT_NAME}"
  
  !insertmacro MUI_PAGE_STARTMENU Application $StartMenuFolder
  
  !insertmacro MUI_PAGE_INSTFILES
  
  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES

;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "English"

;--------------------------------
;Installer Sections

Section "Install"

  SetOutPath "$INSTDIR"

  ;ADD YOUR OWN FILES HERE...
  CreateDirectory "$INSTDIR"
  ;CreateDirectory "$INSTDIR\resources\MPMexports\"
  ;CreateDirectory "$INSTDIR\resources\Namerun\"
  ;CreateDirectory "$INSTDIR\resources\Reports\"
  ;CreateDirectory "$INSTDIR\resources\"
  ;CreateDirectory "$INSTDIR\user guide\"
  
  ;COPY ALL THE FILES THAT NEED TO BE INCLUDED INTO INSTALLER
  file /r "C:\Documents and Settings\G73666\My Documents\workspace\TimeSheet\src\dist\*.*"
  SetOutPath "$INSTDIR\"
  file "C:\Documents and Settings\G73666\My Documents\workspace\TimeSheet\src\msvcp90.dll"
  SetOutPath "$INSTDIR\"
;  file "System.Data.SQLite.dll"
;  SetOutPath "$INSTDIR\resources\databases"
  
 ReadRegStr $R0 HKCU "Software\${SHORT_NAME}" "FileVersion"
 StrCpy $RegistryVersion $R0
 
;     ${If} $RegistryVersion <> ""
;        ${VersionCompare} $RegistryVersion ${version} $R0
;         ${If} $R0 = 2
;              MessageBox MB_YESNO|MB_ICONEXCLAMATION "Overwrite the current ${SHORT_NAME}.db?" IDYES overwrite IDNO backup
;                overwrite: file "resources\databases\${SHORT_NAME}.db"
;                Goto end
;                backup: ${GetTime} "" "LS" $0 $1 $2 $3 $4 $5 $6
;                RENAME "$INSTDIR\resources\databases\${SHORT_NAME}.db" "$INSTDIR\resources\databases\${SHORT_NAME}.$2$1$0$4$5$6.bak" 
;                MessageBox MB_OK "The ${SHORT_NAME}.db has been backed up to ${SHORT_NAME}.$2$1$0$4$5$6.bak"
;                Goto end
;                end:
;         ${EndIf}
;    ${Else}
;    file "resources\databases\${SHORT_NAME}.db"
;    ${EndIf}
;
;  
;  file "resources\databases\WGS84.db"
  
/*
  SetOutPath "$INSTDIR\user guide\"
  file /r /x "MPET" /x "MPET Tutorial.html" "user guide\*.*"
  SetOutPath "$INSTDIR\resources\simulations\6dof_BAMS\"
  file "resources\simulations\6dof_BAMS\*.*"  
  SetOutPath "$INSTDIR\resources\simulations\6dof_Block2\"
  file "resources\simulations\6dof_Block2\*.*"  
  SetOutPath "$INSTDIR\resources\simulations\6dof_Block10\"
  file "resources\simulations\6dof_Block10\*.*" 
  SetOutPath "$INSTDIR\resources\simulations\6dof_Block20\"
  file "resources\simulations\6dof_Block20\*.*" 
  SetOutPath "$INSTDIR\resources\simulations\6dof_Block40\"
  file "resources\simulations\6dof_Block40\*.*"
  SetOutPath "$INSTDIR\resources\simulations\6dof_Eurohawk\"
  file "resources\simulations\6dof_Eurohawk\*.*" 
  SetOutPath "$INSTDIR\resources\simulations\6dof_Eurohawk_L1_2\"
  file "resources\simulations\6dof_Eurohawk_L1_2\*.*" 
  SetOutPath "$INSTDIR\resources\images\"
  file "resources\images\*.*" 
   SetOutPath "$INSTDIR\resources\images\wind radar\"
  file "resources\images\wind radar\*.*" 
  SetOutPath "$INSTDIR\resources\tools\Mp_Val_block_10\"
  file "resources\tools\Mp_Val_block_10\*.*" 
  SetOutPath "$INSTDIR\resources\tools\Mp_Val_block_20\"
  file "resources\tools\Mp_Val_block_20\*.*" 
  */
  
  SetOutPath "$INSTDIR"
  
;  AccessControl::GrantOnFile "$INSTDIR" "(S-1-5-32-545)" "FullAccess"
  
  ;Store installation folder
  WriteRegStr HKCU "Software\${SHORT_NAME}\" "FileVersion" ${version}
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\" "UninstallString" "$PROGRAMFILES\${SHORT_NAME}\Uninstall.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\" "DisplayName" "${SHORT_NAME} v${version}"
;  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\" "DisplayIcon" "$INSTDIR\${FULL_NAME}.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\" "DisplayIcon" "$INSTDIR\driver.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\" "DisplayVersion" "${version}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\" "Publisher" "Northrop Grumman Corporation<joselle.abagat@ngc.com>"
  
  ;Create uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  
  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
    
    ;Create shortcuts
    CreateDirectory "$SMPROGRAMS\$StartMenuFolder"
    CreateShortCut "$SMPROGRAMS\$StartMenuFolder\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
;    CreateShortCut "$SMPROGRAMS\$StartMenuFolder\${SHORT_NAME}.lnk" "$INSTDIR\${FULL_NAME}.exe" "-Xms1024m -Xmx1024m" "$INSTDIR\${FULL_NAME}.exe"
;    CreateShortCut "$DESKTOP\${SHORT_NAME}.lnk" "$INSTDIR\${FULL_NAME}.exe" ""
    CreateShortCut "$SMPROGRAMS\$StartMenuFolder\${SHORT_NAME}.lnk" "$INSTDIR\driver.exe" "$INSTDIR\driver.exe"
    CreateShortCut "$DESKTOP\${SHORT_NAME}.lnk" "$INSTDIR\driver.exe" ""
  
  !insertmacro MUI_STARTMENU_WRITE_END
  
  SectionEnd

;--------------------------------
;Descriptions

  ;Language strings
;  LangString DESC_SecDummy ${LANG_ENGLISH} "A test section."

  ;Assign language strings to sections
;  !insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
;    !insertmacro MUI_DESCRIPTION_TEXT ${SecDummy} $(DESC_SecDummy)
;  !insertmacro MUI_FUNCTION_DESCRIPTION_END
 
;--------------------------------
;Uninstaller Section

Section "Uninstall"

  ;ADD YOUR OWN FILES HERE...
/*  DELETE "$INSTDIR\resources\simulations\6dof_BAMS\*.*"
  DELETE "$INSTDIR\resources\simulations\6dof_Block2\*.*"
  DELETE "$INSTDIR\resources\simulations\6dof_Block10\*.*"
  DELETE "$INSTDIR\resources\simulations\6dof_Block20\*.*"
  DELETE "$INSTDIR\resources\simulations\6dof_Block40\*.*"
  DELETE "$INSTDIR\resources\simulations\6dof_Eurohawk\*.*"
  DELETE "$INSTDIR\resources\simulations\6dof_Eurohawk_L1_2\*.*"
  DELETE "$INSTDIR\resources\images\wind radar\*.*"
  DELETE "$INSTDIR\resources\images\*.*"
  DELETE "$INSTDIR\resources\tools\Mp_Val_block_10\*.*"
  DELETE "$INSTDIR\resources\tools\Mp_Val_block_20\*.*"
  DELETE "$INSTDIR\user guide\*.*"*/
  
  ;DELETE "$INSTDIR\src\*.*"

/*  RMDir "$INSTDIR\resources\simulations\6dof_BAMS\"
  RMDir "$INSTDIR\resources\simulations\6dof_Block2\"
  RMDir "$INSTDIR\resources\simulations\6dof_Block10\"
  RMDir "$INSTDIR\resources\simulations\6dof_Block20\"
  RMDir "$INSTDIR\resources\simulations\6dof_Block40\"
  RMDir "$INSTDIR\resources\simulations\6dof_Eurohawk\"
  RMDir "$INSTDIR\resources\simulations\6dof_Eurohawk_L1_2\"
  RMDir "$INSTDIR\resources\images\wind radar\"
  RMDir "$INSTDIR\resources\images\"
  RMDir "$INSTDIR\resources\tools\Mp_Val_block_10\"
  RMDir "$INSTDIR\resources\tools\Mp_Val_block_20\"
  RMDir "$INSTDIR\resources\tools\"
  RMDir "$INSTDIR\resources\simulations\"
  RMDir "$INSTDIR\"
  RMDir /r "$INSTDIR\user guide" */
  
  ;RMDir "$INSTDIR\src\"
  
  DELETE "$INSTDIR\*.*"
  RMDir "$INSTDIR\"
;  DELETE "$INSTDIR\Uninstall.exe"
;  DELETE "$INSTDIR\ICSharpCode.SharpZipLib.dll"
;  DELETE "$INSTDIR\System.Data.SQLite.dll"
;  DELETE "$INSTDIR\resources\databases\WGS84.db"
  
;  ${GetTime} "" "LS" $0 $1 $2 $3 $4 $5 $6
;  MessageBox MB_YESNO|MB_ICONEXCLAMATION "Backup the local ${SHORT_NAME}.db file? Selecting NO will delete all previous backups in the database subdirectory." IDYES yes IDNO no
;  yes: RENAME "$INSTDIR\resources\databases\${SHORT_NAME}.db" "$INSTDIR\resources\databases\${SHORT_NAME}.$2$1$0$4$5$6.bak" 
;  Goto continue
;  
;  no: DELETE "$INSTDIR\resources\databases\*.*" 
;      RMDIR "$INSTDIR\resources\databases"
;      RMDIR "$INSTDIR\resources"
;  Goto continue
  
;  continue:
;  RMDir "$INSTDIR"
  
  !insertmacro MUI_STARTMENU_GETFOLDER Application $StartMenuFolder
    
  Delete "$SMPROGRAMS\$StartMenuFolder\Uninstall.lnk"
  Delete "$SMPROGRAMS\$StartMenuFolder\${SHORT_NAME}.lnk"
  RMDir "$SMPROGRAMS\$StartMenuFolder"
  
  Delete "$DESKTOP\${SHORT_NAME}.lnk"
  
  DeleteRegKey HKCU "Software\${SHORT_NAME}"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\"

SectionEnd

Function .onInit

 
 ReadRegStr $R0 HKCU "Software\${SHORT_NAME}" "FileVersion"
 Strcpy $RegistryVersion $R0
 
 ${If} $RegistryVersion <> ""
    ${VersionCompare} $RegistryVersion ${version} $R0
 
    ${IF} $R0 = 0
    ${OrIf} $R0 = 1     
          MessageBox MB_YESNO|MB_ICONEXCLAMATION "${SHORT_NAME} $RegistryVersion is already installed. $\n$\nOverwrite ${SHORT_NAME} $RegistryVersion with ${version}?" IDYES uninstall
          Abort
         
        ;Run the uninstaller
        uninstall:
          ClearErrors
           Exec $INSTDIR\Uninstall.exe ; instead of the ExecWait line
         
          IfErrors no_remove_uninstaller done
            ;You can either use Delete /REBOOTOK in the uninstaller or add some code
            ;You can either use Delete /REBOOTOK in the uninstaller or add some code
            ;here to remove the uninstaller. Use a registry key to check
            ;whether the user has chosen to uninstall. If you are using an uninstaller
            ;components page, make sure all sections are uninstalled.
          no_remove_uninstaller:
         done:
     ${ElseIf} $R0 = 2
        MessageBox MB_OK "${SHORT_NAME} $RegistryVersion is currently installed and will be updated to ${version}"
     ${EndIf}
 ${EndIf}
 
FunctionEnd