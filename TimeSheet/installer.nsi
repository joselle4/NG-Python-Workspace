;--------------------------------
;Created on September 17, 2012
;
;@authors: Joselle Abagat, Jesus Medrano, Chris Hoang
;--------------------------------

!include "FileFunc.nsh"
!include "WordFunc.nsh"
!include "LogicLib.nsh"
!include "WinMessages.nsh"
!include "nsProcess.nsh"
!include "Sections.nsh"

;!include "version.txt"
!define version "1.0.2.0"

;NSIS Modern User Interface
;Start Menu Folder Selection Example Script

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
  OutFile "${SHORT_NAME} v${version}.exe"

  ;Default installation folder
  InstallDir "$Desktop\${SHORT_NAME}"
  
  ;Get installation folder from registry if available
  InstallDirRegKey HKCU "Software\${SHORT_NAME}\" "Install Location"

  ;Request application privileges for Windows Vista
  RequestExecutionLevel user

    VIProductVersion "${version}"
    VIAddVersionKey ProductName "${SHORT_NAME}"
    VIAddVersionKey CompanyName "Northrop Grumman Corporation"
    VIAddVersionKey FileDescription "${SHORT_NAME} is a tool to analyze CATS charges"
    VIAddVersionKey LegalCopyright "� 2012 Northrop Grumman Corporation "
    VIAddVersionKey FileVersion "${version}"
    VIAddVersionKey ProductVersion "${version}"

    CRCCheck force
    BrandingText "${FULL_NAME} v${version}"
  
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

  ;CREATE THE FOLDERS YOU WANT TO BE INCLUDED IN THE INSTALLER/ADD YOUR OWN FILES HERE...
  CreateDirectory "$INSTDIR"
  ;CreateDirectory "$INSTDIR\user guide\"
  
  ;COPY ALL THE FILES THAT NEED TO BE INCLUDED INTO INSTALLER
  file /r "C:\Documents and Settings\G73666\My Documents\workspace\TimeSheet\src\dist\*.*" ; /r = recursive
  SetOutPath "$INSTDIR\"
  file "C:\Documents and Settings\G73666\My Documents\workspace\TimeSheet\src\msvcp90.dll"
  SetOutPath "$INSTDIR\"
  file "C:\Documents and Settings\G73666\My Documents\workspace\TimeSheet\src\TSA.ico"
  SetOutPath "$INSTDIR\"
  file "C:\Documents and Settings\G73666\My Documents\workspace\TimeSheet\src\clock.ico"
  SetOutPath "$INSTDIR\"
  file "C:\Documents and Settings\G73666\My Documents\workspace\TimeSheet\src\gh2.jpg"
  SetOutPath "$INSTDIR\"
  file "C:\Documents and Settings\G73666\My Documents\workspace\TimeSheet\src\gradient.jpg"
  SetOutPath "$INSTDIR\"
  
  
 ReadRegStr $R0 HKCU "Software\${SHORT_NAME}" "FileVersion"
 StrCpy $RegistryVersion $R0
  
 SetOutPath "$INSTDIR"
  
;  AccessControl::GrantOnFile "$INSTDIR" "(S-1-5-32-545)" "FullAccess"
  
  ;Store installation folder
  WriteRegStr HKCU "Software\${SHORT_NAME}\" "FileVersion" ${version}
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\" "UninstallString" "$PROGRAMFILES\${SHORT_NAME}\Uninstall.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\" "DisplayName" "${SHORT_NAME} v${version}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\" "DisplayIcon" "$INSTDIR\driver.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\" "DisplayVersion" "${version}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\" "Publisher" "Northrop Grumman Corporation <joselle.abagat@ngc.com>"
  
  ;Create uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  
  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
    
    ;Create shortcuts
    CreateDirectory "$SMPROGRAMS\$StartMenuFolder"
    CreateShortCut "$SMPROGRAMS\$StartMenuFolder\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
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
    
  ;ADD YOUR OWN FOLDERS/FILES YOU WANT REMOVED HERE...
  DELETE "$INSTDIR\*.*"
  RMDir "$INSTDIR\"

  !insertmacro MUI_STARTMENU_GETFOLDER Application $StartMenuFolder
    
  Delete "$SMPROGRAMS\$StartMenuFolder\Uninstall.lnk"
  Delete "$SMPROGRAMS\$StartMenuFolder\${SHORT_NAME}.lnk"
  RMDir "$SMPROGRAMS\$StartMenuFolder"
  
  Delete "$DESKTOP\${SHORT_NAME}.lnk"
  
  DeleteRegKey HKCU "Software\${SHORT_NAME}"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${SHORT_NAME}\"

SectionEnd

Function .onInit

 ;CHECK IF ANY OF THE FILES IN INSTDIR ARE OPEN; IF SO, CLOSE
 ${nsProcess::FindProcess} "Driver.exe" $R0
 ${nsProcess::KillProcess} "Driver.exe" $R0  
 
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