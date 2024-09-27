; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
;                     COMPILER DIRECTIVES
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
.586                      ; create 32 bit code
.model flat, stdcall      ; 32 bit memory model
option casemap :none      ; Case Sensitive
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
;                     INCLUDE FILES
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
    ;
    ;>>>>>>>>>>>>>> 1. TRAY ICON = Include Shell32 >>>>>>>>>>>>>
    ;include \masm32\include\shell32.inc
    ;includelib  \masm32\lib\Shell32.lib
    ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    include \masm32\include\windows.inc
    include \masm32\INCLUDE\dialogs.inc
    ;include \masm32\include\masm32rt.inc
    ;include \masm32\include\masm32.inc
    ;include     \masm32\include\masm32.inc
    ;includelib  \masm32\lib\masm32.lib
    ;
    ;include     \masm32\include\advapi32.inc
    include \masm32\include\kernel32.inc
    include \masm32\include\user32.inc
    include \masm32\include\advapi32.inc
    include \masm32\include\comdlg32.inc
    include \masm32\include\Comctl32.inc    ;Invoke InitCommonControls
    ;include \masm32\include\gdi32.inc
    ;include \masm32\include\Shell32.inc
    ;include \masm32\include\ole32.inc
    ;include \masm32\include\version.inc
    ;
    include     \masm32\include\gdi32.inc
    include     \masm32\include\gdiplus.inc
    include     \masm32\include\msimg32.inc
    include     \masm32\include\ole32.inc
    include     \masm32\include\oleaut32.inc
    include     \masm32\include\Shell32.inc
    include     \masm32\include\version.inc
    include     \masm32\include\winmm.inc
    ;
    ;includelib  \masm32\lib\advapi32.lib
    includelib \masm32\lib\kernel32.lib
    includelib \masm32\lib\user32.lib
    includelib \masm32\lib\advapi32.lib
    includelib \masm32\lib\comdlg32.lib
    includelib \masm32\lib\Comctl32.lib     ;Invoke InitCommonControls
    ;includelib \masm32\lib\gdi32.lib
    ;includelib \masm32\lib\Shell32.lib
    ;includelib \masm32\lib\ole32.lib
    ;includelib \masm32\lib\version.lib
    ;
    includelib \masm32\lib\gdi32.lib
    includelib  \masm32\lib\gdiplus.lib
    includelib  \masm32\lib\msimg32.lib
    includelib  \masm32\lib\ole32.lib
    includelib  \masm32\lib\oleaut32.lib
    includelib  \masm32\lib\Shell32.lib
    includelib  \masm32\lib\version.lib
    includelib  \masm32\lib\winmm.lib
    includelib \masm32\lib\version.lib
    ;
    include macros.inc
    ;
    ; -------------------------------------------
    ; include the image library and include file
    ; -------------------------------------------
    include image.inc
    includelib image.lib

    ; -------------------------------------------
    ; Next Icludes necesary for ImaginA.asm
    ; -------------------------------------------
    ;include \masm32\include\windows.inc
    ;include \masm32\include\dialogs.inc
    ;include \masm32\include\user32.inc
    ;include \masm32\include\kernel32.inc
    ;include \masm32\include\Gdi32.inc
    ;
    ;include \masm32\macros\macros.asm
    ;
    ;includelib \masm32\lib\user32.lib
    ;includelib \masm32\lib\kernel32.lib
    ;includelib \masm32\lib\Gdi32.lib
    ;include .\ImaginA\ImaginA.asm
    Include     \masm32\CarabeZ\AA_Controls\AAGrafics.asm     ;needs msimg32.inc,oleaut32.inc,ole32.inc,gdiplus.inc
    Include     \masm32\CarabeZ\AA_Controls\AAControls.asm    ;needs winmm.inc,comctl32.inc,kernel32.inc
    Include     \masm32\CarabeZ\AA_ImaginA+\ImaginA+.inc
    Includelib  \masm32\CarabeZ\AA_ImaginA+\ImaginA+.lib
    Include     CapturA+.inc ;CapturA+.dll Hooks.

    ;include  .\Service\Service.asm
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;GLOBAL FUNCTIONS
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
    ;wsprintfA           PROTO C :DWORD,:VARARG
    ;wsprintf            equ <wsprintfA>
    ;
    RunCmdAtMainEntry           PROTO ;Result NULL if NOT instance running, else result is hWin of instance running.
    WinMainProc                 PROTO :DWORD,:DWORD,:DWORD,:DWORD
    InitDialogWinMain           PROTO :DWORD ;(WinHandle) Setup initial contents of ListControl, create Toolstips, popupmenus, etc.
    GetThemeFonts               PROTO
    SetTheme                    PROTO :DWORD,:DWORD ;(WinHandle), pointer to FONTDATA<>
    SelectPage                  PROTO :DWORD ;dPageID, Select Active Page: Shows controls from page ID and hides others
    ShowPage                    PROTO :DWORD,:DWORD ;dPageID, nCmdShow
    UpdateActivityLog           PROTO :DWORD,:DWORD,:BYTE,:BYTE ;pActivityTitle,pActivityText,bNewLinesAfterTitle,bNewLinesAfterText (0,0 = Put szActivityLogText)
    ;
    ;
    RegionDlg                   PROTO :DWORD,:DWORD
    RegionDlgProc               PROTO :DWORD,:DWORD,:DWORD,:DWORD
    ;
    ;
    IniManager                  PROTO :DWORD,:BYTE ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
    ValidateIniValue            PROTO :DWORD,:DWORD ;pSource,dSize
    GetRegistryMouseMenu        PROTO :DWORD,:DWORD,:DWORD ;szMouseHoverTime, szMenuShowDelay,szDoubleClickSpeed
    UpdateRegistryRun           PROTO :DWORD ;Create or Delete Regitry entry
    UpdateRegistryPrintScreen   PROTO :DWORD ;NewValue for dPrintScreenKeyForSnipping, return Old value
    ServiceManager              PROTO :DWORD,:DWORD,:DWORD,:BYTE ;(ServiceName),(ServiceDesc),(ServicePath),(INSTALL,REMOVE,ISINSTALLED,ISRUNNING)
    ErrorManager                PROTO :DWORD ;(Title)
    ;
    TakeScreenShot              PROTO :BYTE,:BYTE              ;bLocalCaptureWindow:BYTE,bSendArchiveTo:BYTE
    StartAutoShotLoop           PROTO :BYTE,:BYTE,:BYTE,:BYTE  ;                         bSendArchiveTo:BYTE,bMinutes:BYTE,bSeconds:BYTE,bOnlyOneTime:BYTE
    AutoShotLoop                PROTO :DWORD
    StopAutoShotLoop            PROTO
    ShowMainWindow              PROTO :BYTE,:DWORD ;bShowWindow Hide=FALSE/show=TRUE,dChangeWinMainExStyle NULL/NewSlyle
    ;
    SolveMacros                 PROTO :DWORD,:DWORD                         ;lpSourceStr, lpTarget
    ReplaceMacros               PROTO :DWORD,:DWORD,:DWORD                  ;lpSourceStr, lpPattern, lpTarget
    PasteMacro                  PROTO :DWORD,:DWORD,:DWORD                  ;PopUpSelection,EditHandle,EditID
    SubWndProc                  PROTO :DWORD,:DWORD,:DWORD,:DWORD
    ;
    ConvertStaticToHyperlink    PROTO :DWORD,:DWORD
    _HyperlinkParentProc        PROTO :DWORD,:DWORD,:DWORD,:DWORD
    _HyperlinkProc              PROTO :DWORD,:DWORD,:DWORD,:DWORD
    ;
    Paint_Proc                  PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
    ;
    StrLen                      PROTO :DWORD
    InString                    PROTO :DWORD,:DWORD,:DWORD
    FillBuffer                  PROTO :DWORD,:DWORD,:BYTE
    GetFileName                 PROTO :DWORD,:DWORD,:DWORD
    GetCL                       PROTO :DWORD,:DWORD
    a2dw                        PROTO :DWORD
    FindApplicationProcessID    PROTO :DWORD,:DWORD,:DWORD                  ;pszApplicationPath,dWhatToDoIfFound,pOutWinHandle
    EnumWindowsProc             PROTO :DWORD,:DWORD                         ;hwinparent,lparam=userdefinedparam
    ForceWindowToTop            PROTO :DWORD
    FindTopWindowFromProcessID  PROTO :DWORD,:DWORD                 ;dProcessID,dWhatToDoIfFound
    GetProcessMainThreadID      PROTO :DWORD ; hProcessID
    SendKey                     PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD, ;SendKey(WORD wVk, int nSleep, BOOL bExtendedkey, BOOL bDown, BOOL bUp)
    ;
.const
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;EQUATES
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
    ;
    RUNTIME_TARGET_TYPE_AUTOSHOT        equ 0
    RUNTIME_TARGET_TYPE_MONITOR         equ 1
    RUNTIME_TARGET_TYPE_ACTIVE_WINDOWS  equ 2
    RUNTIME_TARGET_TYPE_DESKTOP         equ 3
    RUNTIME_TARGET_TYPE_REGION          equ 4
    ;
    RUNTIME_CAPTURE_CURSOR_DISABLE      equ 1
    RUNTIME_CAPTURE_CURSOR_SYSTEM_ARROW equ 2
    RUNTIME_CAPTURE_CURSOR_RED_ARROW    equ 3
    ;
    RUNTIME_IMAGE_MODE_COLOR            equ 1
    RUNTIME_IMAGE_MODE_GRAY_SCALE       equ 2
    ;
    RUNTIME_CAPTURE_MODE_AUTOSHOT       equ 1
    RUNTIME_CAPTURE_MODE_MANUAL         equ 2
    RUNTIME_CAPTURE_MODE_CLICKS         equ 3
    ;
    RUNTIME_CAPTURE_SENT_TO_NONE        equ 0
    RUNTIME_CAPTURE_SENT_TO_CLIPBOARD   equ 1
    RUNTIME_CAPTURE_SENT_TO_OPEN        equ 2
    RUNTIME_CAPTURE_SENT_TO_MAIL        equ 3
    RUNTIME_CAPTURE_SENT_TO_LOCATION    equ 4
    ;
    INSTALL                             equ 0
    REMOVE                              equ 1
    ISINSTALLED                         equ 2
    ISRUNNING                           equ 3
    ;
    MAX_CLICK_DELAY                     equ 4000 ;Type in a number between 0 to 4000 (400 is default)
    ;
    ;
    IDC_LIST_GROUP_OPTIONS              equ 001
    IDC_LIST_ITEM_BUTTON1               equ 002
    IDC_LIST_ITEM_BUTTON2               equ 003
    IDC_LIST_ITEM_BUTTON3               equ 004
    IDC_LIST_ITEM_BUTTON4               equ 005
    IDC_LIST_ITEM_BUTTON5               equ 006
    IDC_LIST_ITEM_BUTTON6               equ 007
    IDC_LIST_ITEM_BUTTON7               equ 008
    IDC_LIST_ITEM_BUTTON8               equ 009
    ;
    IDC_SHEET_TAB_CONTROL               equ 010
    ;
    IDC_PROMPT_TITLE_CAPTURE            equ 011
    IDC_PROMPT_SUBTITLE_CAPTURE         equ 012
    IDC_PROMPT_CAPTURE_TARGET           equ 013
    IDC_DROP_CAPTURE_TARGET             equ 014
    IDC_BTN_DIALOG_CAPTURE_REGION       equ 015
    IDC_PROMPT_CAPTURE_CURSOR           equ 016
    IDC_DROP_CAPTURE_CURSOR             equ 017
    IDC_PROMPT_COLOR_IMAGE_MODE         equ 018
    IDC_DROP_COLOR_IMAGE_MODE           equ 019
    IDC_PROMPT_CAPTURE_MODE             equ 020
    IDC_DROP_CAPTURE_MODE               equ 021
    IDC_PROMPT_TIMELAPSE                equ 022
    IDC_EDITBOX_TIMELAPSE_MIN           equ 023
    IDC_SPINBOX_TIMELAPSE_MIN           equ 024
    IDC_EDITBOX_TIMELAPSE_SEC           equ 025
    IDC_SPINBOX_TIMELAPSE_SEC           equ 026
    IDC_PROMPT_START_AUTOSHOT_AT_INIT   equ 027
    IDC_CHK_START_AUTOSHOT_AT_INIT      equ 028
    IDC_PROMPT_CLICK_DELAY              equ 029
    IDC_EDITBOX_CLICK_DELAY             equ 030
    IDC_SPINBOX_CLICK_DELAY             equ 031

    IDC_BTN_AUTOSHOT                    equ 032
    IDC_BTN_SAVE                        equ 033
    IDC_BTN_SEND_TO_LOCATION            equ 034
    IDC_BTN_SEND_TO_EMAIL               equ 035
    IDC_BTN_SEND_TO_CLIPBOARD           equ 036
    IDC_BTN_HIDE                        equ 037
    ;
    IDC_PROMPT_TITLE_FILE               equ 038
    IDC_PROMPT_COPY_PATH_TO_CLIPBOARD   equ 039
    IDC_CHK_COPY_PATH_TO_CLIPBOARD      equ 040
    IDC_PROMPT_SUBTITLE_FILE            equ 041
    IDC_PROMPT_TARGET_FILE_FORMAT       equ 042
    IDC_DROP_TARGET_FILE_FORMAT         equ 043
    IDC_PROMPT_ARCHIVE_PATH             equ 044
    IDC_EDITBOX_ARCHIVE_PATH            equ 045
    IDC_BTN_DIALOG_ARCHIVE_PATH         equ 046
    IDC_PROMPT_COUNTER                  equ 047
    IDC_EDITBOX_COUNTER                 equ 048
    IDC_SPINBOX_COUNTER                 equ 049
    IDC_PROMPT_IMAGE_QUALITY            equ 050
    IDC_EDITBOX_IMAGE_QUALITY           equ 051
    IDC_SPINBOX_IMAGE_QUALITY           equ 052
    IDC_PROMPT_IMAGE_ZOOM               equ 053
    IDC_EDITBOX_IMAGE_ZOOM              equ 054
    IDC_SPINBOX_IMAGE_ZOOM              equ 055
    IDC_PROMPT_JPG_COMMENTS             equ 056
    IDC_EDITBOX_JPG_COMMENTS            equ 057
    ;
    IDC_PROMPT_TITLE_APP                equ 058
    IDC_PROMPT_SUBTITLE_APP             equ 059
    IDC_PROMPT_WIN_STARTUP_RUN          equ 060
    IDC_CHK_WIN_STARTUP_RUN             equ 061
    IDC_PROMPT_HIDE_WIN_WHEN_OPENING    equ 062
    IDC_CHK_HIDE_WIN_WHEN_OPENING       equ 063
    IDC_PROMPT_PLAY_AUDIBLE_SHOT        equ 064
    IDC_CHK_PLAY_AUDIBLE_SHOT           equ 065
    IDC_PROMPT_OPEN_LAST_CAPTURE_PATH   equ 066
    IDC_CHK_OPEN_LAST_CAPTURE_PATH      equ 067
    IDC_PROMPT_COPY_SCREENSHOT_CLIPBOARD equ 068
    IDC_CHK_COPY_SCREENSHOT_CLIPBOARD   equ 069
    IDC_PROMPT_SHOW_NOTIFICATION_ICON   equ 070
    IDC_CHK_SHOW_NOTIFICATION_ICON      equ 071
    IDC_PROMPT_DISPLAY_NOTIFICATION_CAPTURE equ 072
    IDC_CHK_DISPLAY_NOTIFICATION_CAPTURE equ 073
    ;
    IDC_PROMPT_HOTKEYS_TITLE            equ 074
    IDC_PROMPT_HOTKEYS_SUBTITLE         equ 075
    IDC_PROMPT_HOTKEYS_SHOW_ME          equ 076
    IDC_EDITBOX_HOTKEYS_SHOW_ME         equ 077
    IDC_PROMPT_HOTKEYS_SHOW_ME_LEFT     equ 078
    IDC_EDITBOX_HOTKEYS_SHOW_ME_LEFT    equ 079
    IDC_PROMPT_HOTKEYS_CAPTURE_TARGET   equ 080
    IDC_EDITBOX_HOTKEYS_CAPTURE_TARGET  equ 081
    IDC_PROMPT_HOTKEYS_FOR_EDIT_TARGET  equ 082
    IDC_EDITBOX_HOTKEYS_FOR_EDIT_TARGET equ 083
    IDC_PROMPT_HOTKEYS_CAPTURE_TOPWIN   equ 084
    IDC_EDITBOX_HOTKEYS_CAPTURE_TOPWIN  equ 085
    IDC_PROMPT_HOTKEYS_FOR_EDIT_TOPWIN  equ 086
    IDC_EDITBOX_HOTKEYS_FOR_EDIT_TOPWIN equ 087
    IDC_PROMPT_HOTKEYS_EMAIL_TARGET     equ 088
    IDC_EDITBOX_HOTKEYS_EMAIL_TARGET    equ 089
    IDC_PROMPT_HOTKEYS_EMAIL_TOPWIN     equ 090
    IDC_EDITBOX_HOTKEYS_EMAIL_TOPWIN    equ 091
    ;
    IDC_PROMPT_ACTIVITY_TITLE           equ 092
    IDC_PROMPT_ACTIVITY_SUBTITLE        equ 093
    IDC_PROMPT_ACTIVITY_LOG             equ 094
    IDC_EDITBOX_ACTIVITY_LOG            equ 095
    ;
    IDC_PROMPT_ABOUT_TITLE              equ 096
    IDC_PROMPT_ABOUT_SUBTITLE           equ 097
    IDC_PROMPT_ABOUT_LOG                equ 098
    IDC_EDITBOX_ABOUT_LOG               equ 099
    ;
    IDC_MENU_BAR                        equ 1000
    ;
    ;>>>>>>>>>>>>>> 2. TRAY ICON = PopUp Menu Control Equates >>>>>>>>>>>>>
    IDM_TRAY_SHELLNOTIFY_MSG            equ 101
    IDM_TRAY_SHOW_APPLICATION           equ 102
    IDM_TRAY_CAPTURE_SENT_TO_CLIPBOARD  equ 103
    IDM_TRAY_CAPTURE_SENT_TO_LOCATION   equ 104
    IDM_TRAY_CAPTURE_SENT_TO_MAIL       equ 105
    IDM_TRAY_CAPTURE_SENT_TO_OPEN       equ 106
    IDM_TRAY_QUIT_APPLICATION           equ 107
    ;
    ;
    HOTKEY_SHIFT_CTRL_ALT_PRINTSCREEN   equ 108
    HOTKEY_SHIFT_CTRL_ALT_S             equ 109
    HOTKEY_PRINTSCREEN                  equ 110
    HOTKEY_ALT_PRINTSCREEN              equ 111
    HOTKEY_CTRL_PRINTSCREEN             equ 112
    HOTKEY_CTRL_ALT_PRINTSCREEN         equ 113
    HOTKEY_SHIFT_CTRL_E                 equ 114
    HOTKEY_SHIFT_CTRL_ALT_E             equ 115
    WM_SHELLNOTIFY                      equ WM_USER+1913
    COLOR_POLLO                         equ 0E1FFFFH
    ;
    ;WM_SHELLNOTIFY             equ WM_USER+10005
    ;IDI_TRAY                   equ 0
    ;IDM_FIRST                  equ 1001
    ;IDM_TRAY_QUIT_APPLICATION                   equ 1002
    NIF_INFO                    equ 10h
    ;NIM_SETVERSION             equ 0 ;Use the Windows 95 behavior=Windows versions prior to Windows 2000
    NOTIFYICON_VERSION          equ 3 ;Window 2000 and later.
    NIIF_NONE                   equ 0h
    NIIF_INFO                   equ 1h
    NIIF_WARNING                equ 2h
    NIIF_ERROR                  equ 3h
    ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ;
    WM_MOUSEHOOK                equ WM_USER+492
    ;
    ;PopUp MenuMacros
    IDM_01                      equ 100
    IDM_02                      equ 101
    IDM_03                      equ 102
    IDM_04                      equ 103
    IDM_05                      equ 104
    IDM_06                      equ 105
    IDM_07                      equ 106
    IDM_08                      equ 107
    IDM_09                      equ 108
    IDM_10                      equ 109
    IDM_11                      equ 110
    IDM_12                      equ 211
    IDM_13                      equ 112
    IDM_14                      equ 113
    IDM_15                      equ 114
    IDM_16                      equ 115
    IDM_17                      equ 116
    IDM_18                      equ 117
    IDM_19                      equ 118
    IDM_20                      equ 119
    IDM_21                      equ 120
    IDM_22                      equ 121
    IDM_23                      equ 122
    IDM_24                      equ 123
    IDM_25                      equ 125
    ;
    IDGROUP3                    equ 200    ;About Window equates
    IDSTATIC12                  equ 201
    IDC_BTN6                    equ 202
    ;
    IDC_PAGE01                  equ 001
    IDC_PAGE02                  equ 002
    IDC_PAGE03                  equ 003
    IDC_PAGE04                  equ 004
    IDC_PAGE05                  equ 005
    IDC_PAGE06                  equ 006
    IDC_LASTPAGE                equ 006
    ;
.data
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;INITIALIZED DATA SECTION
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
    ;
    szAppTitle                  db "CapturA+ Screen Capture",0
    szAtom                      db "Carabez",0,0;"#1913",0
    szCaptureWindow             db 60 dup('Active Monitor',0,'Active Window',0,0,'DeskTop',0,0,0,0,0,0,0,0,'Region',0,0,0,0,0,0,0,0,0);15
    szCaptureCursor             db 45 dup('Disabled',0,0,0,0,0,0,0,'Default Cursor',0,'Red Arrow',0,0,0,0,0,0);15
    szCaptureMode               db 36 dup('Auto Shot',0,0,0,'Manual Shot',0,'Mause Click',0);12
    szImageMode                 db 22 dup('Color',0,0,0,0,0,0,'Gray Scale',0);11
    szFileType                  db 40 dup('JGP File (*.jpg)',0,0,0,0,'PNG File (*.png)',0,0,0,0);20
    szStoreMode                 db 30 dup('Single Files',0,0,0,'Append to File',0);15
    ;
    ;szNoFilter                 db "Program Files",0,"*.exe",0,"All files",0,"*.*",0,0
    szCOMPUTERNAME              db "%COMPUTERNAME%",0         ;Computer name of the current system
    szCOUNTER                   db "%COUNTER%",0              ;SnapShot Counter
    szDATE                      db "%DATE%",0                 ;Current date (in the format YYYYMMDD)
    szDAY                       db "%DAY%",0                  ;Day of the month (1-31)
    szDRIVE                     db "%DRIVE%",0                ;Drive that is redirected to ShareAddress (D  db  C  db  etc)
    szERRORCODE                 db "%ERRORCODE%",0            ;Last Error Code Returned by  GetLastError
    szERRORTEXT                 db "%ERRORTEXT%",0            ;Last Error Text Returned by  GetLastError
    szHOUR                      db "%HOUR%",0                 ;LocalTime Hour 0-24
    szMINUTE                    db "%MINUTE%",0               ;LocalTime Hour Minute 00-59
    szMONTH                     db "%MONTH%",0                ;Month of the Year (1-12)
    szPROGRAMDIR                db "%PROGRAMDIR%",0           ;Path of the Program Working directory
    szPROGRAMFILES              db "%PROGRAMFILES%",0         ;Path of the Program Files Directory
    szPROGRAMNAME               db "%PROGRAMNAME%",0          ;Name of the current executable file  usually = MapDrive
    szSECOND                    db "%SECOND%",0               ;LocalTime Second 00-59
    szSHARE                     db "%SHARE%",0                ;Full Share Address
    szSYSTEMDIR                 db "%SYSTEMDIR%",0            ;Path of the Windows system directory
    szSYSTEMDRIVE               db "%SYSTEMDRIVE%",0          ;Default Root Drive usually = C:\
    szTEMPDIR                   db "%TEMP%",0                 ;Default Temp Directory
    szTIME                      db "%TIME%",0                 ;Current time (in the format HHMMSS)
    szUSERNAME                  db "%USERNAME%",0             ;Current user's Windows user ID
    szWEEKDAY                   db "%WEEKDAY%",0              ;Days since Sunday (1 – 7)
    szWINDIR                    db "%WINDIR%",0               ;Path of the Windows directory
    szYEAR                      db "%YEAR%",0                 ;Current year
    ;
    szformat                    db '%02d',0                 ; To format '1' into '01:'
    szIntFormat                 db "%lu",0                  ;To convert dw to ascci
    szNoFilter                  db "Jpg Files",0,"*.jpg",0,"All files",0,"*.*",0,0
    ;
    ;
    szWebPage                   db 'http://www.carabez.com',0
    PROP_ORIGINAL_FONT          db '_Hyperlink_Original_Font_',0
    PROP_ORIGINAL_PROC          db '_Hyperlink_Original_Proc_',0
    PROP_UNDERLINE_FONT         db '_Hyperlink_Underline_Font_',0
    szFileVersion               db "\\StringFileInfo\\040904E4\\FileVersion",0
    ;
    szErrorTitleB               db "Error.",0
    szErrorTextB                db "Readonly Drive or The ""Archive Path"" is NOT valid.",0
    ;
    ;szClassName                 db "WinClass",0
    ;wc                          WNDCLASSEX  <sizeof WNDCLASSEX,\
    ;                                        CS_HREDRAW or CS_VREDRAW or CS_BYTEALIGNWINDOW,\
    ;                                        offset DlgProc,NULL,NULL,NULL,0,0,COLOR_BTNFACE+1,\
    ;                                        NULL,offset szClassName,0>
    szTOOLTIPS_CLASS            db "tooltips_class32",0
                                ;length for standard tooltip text is 80 characters
    szToolTxt0                  db 9,"Press CTRL+SPACEBAR",10,"to show selection PopUp Paste-In-Macros menu",0
    szToolTxt1                  db "Minutes",0
    szToolTxt2                  db "Seconds",0
    szToolTxt3                  db "JPEG encoder quality parameter:",10,10
                                db 9,"Range between 1% and 100%",10,10
                                db "As a general benchmark:",10,10
                                db 9,"90% JPEG quality gives a very high-quality image while ",10,10,9,"gaining a significant reduction on the original 100% file size.",10,10
                                db "Note: To get similar PNG archive file size use 85%",0
                                ;
    szToolTxt4                  db "Value for Paste-In-Macro %COUNTER%",10,10,"This value increments after every shot.",0
    ;
    szToolTxt5                  db "Press PRINTSCREEN key to take a screenshot of Target.",10,10
                                db "ALT+PRINTSCREEN to take a quick screenshot of the active window.",10,10
                                db "CTRL+PRINTSCREEN to take a screenshot of Target and open Windows default app for jpeg files.",10,10
                                db "CTRL+ALT+PRINTSCREEN to take a screenshot of active window and open Windows default app for jpeg files.",0
    ;
    szToolTipTxt_Mode           db "Auto Shot: Timelapse Screenshot Recorder",10,10
                                db "Manual Shot: Press PRINTSCREEN key",10,10
                                db "Mouse Click: Screenshot on each left mouse click",0

    szToolTipTxt_Cursor         db "Use a cursor file of your choice",10,10,"by simply dropping it into the application folder",10,10,"and renaming it to cursor.cur",0

    szToolTipTxt_Delay1         db "To capture menu items this value must be ~200 greater than",10
                                db "MenuShowDelay* from HKEY_CURRENT_USER\Control Panel\Desktop = ",0

    szToolTipTxt_Delay2         db 10,10,"To capture tooltip balloon this value must be ~200 greater than",10
                                db "MouseHoverTime* from HKEY_CURRENT_USER\Control Panel\Mouse = ",0
;
    szToolTipTxt_Delay3         db 10,10,"*Range from 1 to 4000, default value 400 (ms).",0

    szToolTipTxt_Zoom           db "To get an increase or decrease image dimensions by percent",10,9,"          scale the image proportionally:",10,10
                                db "50% will half the image dimensions",10,10
                                db "200% will double the image dimensions",0
    ;
    szToolTxt8                  db 9,"HIDES THIS WINDOW",10,10,"Press SHIFT+CTRL+ALT+PRINTSCREEN",10,09,09,"or",10
                                db "Press SHIFT+CTRL+ALT+S ",10," ",10,9,"to unhide Window setting.",0
    ;
    ;
    szExtentionDotExe           db ".exe",0
    szBracketOpen               db " [",0
    szBracketClose              db "]",0
    szQuote                     db '"',0
    ;
    szSystemFontName            db "Segoe UI",0 ;"Segoe UI Variable",0 ;"Segoe UI",0
    szSystemFontNameVar         db "Segoe UI Variable",0 ;"Segoe UI",0
    FONTDATA                    struct
                                    hTitleFont          dd 0
                                    hSubTitleFont       dd 0
                                    hRegularFont        dd 0
                                    hBrushAccent        dd 0 ;#0078D4
                                    hBrushHigh          dd 0 ;BaseHigh color = contrast color
                                    dColorTransparent   dd 0
                                    dColorAccent        dd 0 ;#0078D4
                                    IsRedrawNeeded      dd 0
    FONTDATA                    ends
    LightMode                   FONTDATA<>
    DarkMode                    FONTDATA<>
    dLastSelectedPage           dd 0
    szText_normal               db "Off",0 ;"Activado",0
    szText_push                 db "On",0 ;"Desactivado",0
    szText_StartAutoShot        db "Start Auto Shot",0
    szText_StopAutoShot         db "Stop Auto Shot",0
    ;https://learn.microsoft.com/en-us/windows/apps/design/style/color
    ;
    szHOTKEY_SHIFT_CTRL_ALT_PRINTSCREEN     db "SHIFT+CTRL+ALT+PRINTSCREEN",0
    szHOTKEY_SHIFT_CTRL_ALT_S               db "SHIFT+CTRL+ALT+S",0
    szHOTKEY_PRINTSCREEN                    db "PRINTSCREEN",0
    szHOTKEY_ALT_PRINTSCREEN                db "ALT+PRINTSCREEN",0
    szHOTKEY_CTRL_PRINTSCREEN               db "CTRL+PRINTSCREEN",0
    szHOTKEY_CTRL_ALT_PRINTSCREEN           db "CTRL+ALT+PRINTSCREEN",0
    szHOTKEY_SHIFT_CTRL_E                   db "SHIFT+CTRL+E",0
    szHOTKEY_SHIFT_CTRL_ALT_E               db "SHIFT+CTRL+ALT+E",0
    ;
    szNewLine                               db 13,10,0
    szApplicationHistory                    db "Version 2.2024.09.26",13,10
                                            db "-Fix bug when starting hidden.",13,10,13,10
                                            db "Version 2.2024.09.10",13,10
                                            db "-New look and feel but the same 'what you see is what you get'.",13,10,13,10
                                            db "Version 1.2024.08.30",13,10
                                            db "-New: Added option to capture Active Monitor (the one that have the active windows with focus).",13,10,13,10
                                            db "Version 1.2024.05.05",13,10
                                            db "-New: Added option to delay snap on mouse clicks.",13,10,13,10
                                            db "Version 1.2024.05.04",13,10
                                            db "-CapturA+.exe can run without CapturA+.dll but capturing each mouse click option will be removed.",13,10,13,10
                                            db "Version 1.2024.04.30",13,10
                                            db "-New: Saves and restore HKEY_CURRENT_USER\Control Panel\Keyboard\PrintScreenKeyForSnippingEnabled value on application init and close.",13,10,13,10
                                            db "Version 1.2024.04.25",13,10
                                            db "-New: Added Hotkeys features, vgr. ALT+PRINTSCREEN to capture only the image of the active window.",13,10,13,10
                                            db "Version 1.2024.04.15",13,10
                                            db "-Bug: Fixed Bad initial data for File path.",13,10,13,10
                                            db "Version 1.2019.10.04",13,10
                                            db "-New: Option to capture Cursor.",13,10
                                            db "-New: Option to capture on each Mouse clicks",13,10,13,10
                                            db "Version 1.2019.10.05",13,10
                                            db "-New: Options saved to Register if .ini archive does NOT exists.",13,10
                                            db "-New: Single Intance mode.",13,10
                                            db "-Change: Image Quality = Compression rate decreasing vs quality increasing.",13,10,13,10
                                            db "Version 1.2010.05.07",13,10
                                            db "-First version release.",13,10,13,10
                                            db "Version 0.2010.02.25",13,10
                                            db "-Functional beta release.",0
.data?
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;UNITIALIZED DATA SECTION
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
    szFullPathLocal             db 4096 dup(?)
    szOutputDebugString         db 4096 dup(?)
    hPreviousWin                dd ?
    COLOR_NONE                  dd ?
    hPenPollo                   dd ?
    ;
    dFirstTimeShowed            dd ?
    WinPos                      RECT<?>
    ;
    hInstance                   dd ?
    hCursorRed                  dd ?
    hCursor                     dd ?
    hIconTray                   dd ?
    hIcon                       dd ?
    hIconOnOffAscent            dd ?
    hIconOnOffUnpush            dd ?
    hIconCameraAscent           dd ?
    hIconCameraUnpush           dd ?
    hIconFileAscent             dd ?
    hIconFileUnpush             dd ?
    hIconSettingsAscent         dd ?
    hIconSettingsUnpush         dd ?
    hIconHotkeysAscent          dd ?
    hIconHotkeysUnpush          dd ?
    hIconWarningAscent          dd ?
    hIconWarningUnpush          dd ?
    hIconAboutAscent            dd ?
    hIconAboutUnpush            dd ?
    hIconExitAscent             dd ?
    hIconExitUnpush             dd ?
    hIconSaveAscent             dd ?
    hIconSaveUnpush             dd ?
    hIconAutoshotAscent         dd ?
    hIconAutoshotUnpush         dd ?
    hIconToMailAscent           dd ?
    hIconToMailUnpush           dd ?
    hIconToClipAscent           dd ?
    hIconToClipUnpush           dd ?
    hIconToPathAscent           dd ?
    hIconToPathUnpush           dd ?
    hIconHideAscent             dd ?
    hIconHideUnpush             dd ?

    hBmp                        dd ?
    hWinMain                    dd ?
    hBkbrush                    dd ?
    hSingleInstanceAtom         dd ?
    hWinMainInstanceAtom        dd ?
    bWinMainInstanceHotSendKeys dd ?
    dHandleFlags                dd ?
    lpListProc                  dd ?   ;ListProc
    ;
    bSelectingPage              dd ?
    bHookRunning                dd ?
    bExeIsClosing               dd ?
    bExeCloseIsAllowed          db ?
    bIniArchiveExists           db ?
    ;
    ;TakeScreenShot and AutoShotLoop
    dRunningCapture             dd ?
    dAutoShotLoop_IsRunning    dd ? ;0=Stopped,1=Running,2=Stopping
    dAutoShotLoop_ThreadID     dd ?
    hAutoShotLoop_ThreadHandle dd ?
    ;
    ASL_PARAM                   struct
        bSendArchiveTo              db ?
        bMinutes                    db ?
        bSeconds                    db ?
        bOnlyOneTime                db ?
    ASL_PARAM                   ends
    sAutoShotLoop_Param         ASL_PARAM <?>
   ;
    szProgramPath               db MAX_PATH dup(?)
    szImageFilePath             db MAX_PATH dup(?)
    szProgramName               db MAX_PATH dup(?)
    szAppName                   db MAX_PATH dup(?)
    szQuotedPath                db MAX_PATH+MAX_PATH dup(?)                ;Quoted full Program Path to put on registry
    szIniFile                   db MAX_PATH dup(?)
    szCurFile                   db MAX_PATH dup(?)
    szBuffer                    db MAX_PATH dup(?)
    szVer                       db MAX_PATH dup(?)
    ofn                         OPENFILENAME  <?>                   ;FileDialog structure
    ;
    ;
    ;Macros
    dErrorCode                  dd ?
    szWindowsDir                db MAX_PATH dup(?)
    hMacroPopupMenu             dd ?                     ;Handle of PopUpMenu
    hOwnerPopMenu               dd ?                     ;Active Windows'Handle target of PopUpMenu
    hFocus                      dd ?                     ;To get Handle of control with focus
    ACTIVESIDC                  struct
        hIDC                    dd ?
        hProc                   dd ?
    ACTIVESIDC                  ends
    ActiveIDC                   ACTIVESIDC 2 dup(<?>)
    ;
    hCombo1                     dd ?
    hCombo2                     dd ?
    hCombo3                     dd ?
    hCombo4                     dd ?
    hCombo5                     dd ?
    hCombo6                     dd ?
    hCombo7                     dd ?
    hProgress                   dd ?
    hwndStatus                  dd ?
    ;
    pPointer                    dd ?
    dCounter                    dd ?
    ;

    ;SystemTime                 SYSTEMTIME    <?>
    ;FileTime                   FILETIME      <?>
    ;
    ;AboutDlg Relative
    hOldStaticProc              dd ?
    ;
    ;Version
    dDummy                      dd ?
    pMemory                     dd ?
    dwVersionLen                dd ?
    ;
    ;SnapShot Settings
    dPrintScreenKeyForSnipping  dd ?        ;NewValue for Win10/11 dPrintScreenKeyForSnipping, return Old value
    rcCapture                   RECT <?>    ;Used for capturing Region/Area
    ;
    ;
    ; RUNTIME VARIABLES FOR PAGE "CAPTURE"

    bCaptureWindow              db ?  ;TARGET TYPE: RUNTIME_TARGET_TYPE_MONITOR,RUNTIME_TARGET_TYPE_ACTIVE_WINDOWS,RUNTIME_TARGET_TYPE_DESKTOP,RUNTIME_TARGET_TYPE_REGION
    bCaptureCursor              db ?  ;CAPTURE CURSOR: RUNTIME_CAPTURE_CURSOR_DISABLE,RUNTIME_CAPTURE_CURSOR_SYSTEM_ARROW,RUNTIME_CAPTURE_CURSOR_RED_ARROW
    bImageMode                  db ?  ;IMAGE MODE: RUNTIME_IMAGE_MODE_COLOR, RUNTIME_IMAGE_MODE_GRAY_SCALE
    bCaptureMode                db ?  ;CAPTURE MODE: RUNTIME_CAPTURE_MODE_AUTOSHOT,RUNTIME_CAPTURE_MODE_MANUAL,RUNTIME_CAPTURE_MODE_CLICKS
    dMouseClickSnapDelay        dd ?  ;WHEN ;CAPTURE MODE = RUNTIME_CAPTURE_MODE_CLICKS
    szMouseHoverTime            db MAX_PATH dup(?) ;TOOLTIP INFO TEXT FOR dMouseClickSnapDelay EDIT
    szMenuShowDelay             db MAX_PATH dup(?) ;TOOLTIP INFO TEXT FOR dMouseClickSnapDelay EDIT
    szDoubleClickSpeed          db MAX_PATH dup(?) ;dMouseDoubleClick Speed to Tray popupmenu
    dDoubleClickSpeed           dd ?
    bAutoShot_TimerMinutes      db ?
    bAutoShot_TimerSeconds      db ?
    bAutoShotStartAtInit        db ?
    ;
    dIsWindowStyleTOOLWINDOW    dd ? ;When open Hidden, style is change to TOOLWINDOW to hide Taskbar Application Icon
    ;
    ; RUNTIME VARIABLES FOR PAGE "FILE"
    bCopyFilePathToClipboard    db ?
    dSnapShotCounter            dd ?
    bImageQuality               db ?
    bImageZoom                  db ?
    ;
    ; RUNTIME VARIABLES FOR PAGE "APPLICATION"
    bRunWhenWindowsStarts                   db ?
    bHideApplicationWindowWhenOpening       db ?
    bIsUnhideApplicationWindowHotKeyActive  db ? ;Auxiliar
    dWinMainExStyle                         dd ? ;Auxiliar for hide/show Main Window at init
    bAudibleNotificationAfterCapture        db ?
    bRestoreWindowAfterCapture              db ?
    bOpenLastCaptureLocationPath            db ?
    bCopyScreenShotToClipboard              db ?
    bShowANotificationIcon                  db ?
    bIsShowANotificationIconActive          db ? ;Auxiliar for show/hide Tray Icon
    bDisplayANotificationCapture            db ?

    ;
    ;Target Setting
    bTargetType                 db ?
    bArchiveMode                db ?
    szArchivePath               db MAX_PATH dup(?)
    szJpgComments               db MAX_PATH dup(?)
    ;
    ;>>>>>>>>>>>>>> 3. TRAY ICON = Global Variables >>>>>>>>>>>>>
    ;
    dDoubleClick                dd ?
    NOTIFYICONDATAAA STRUCT
      cbSize                    DWORD      ?
      hwnd                      DWORD      ?
      uID                       DWORD      ?
      uFlags                    DWORD      ?
      uCallbackMessage          DWORD      ?
      hIcon                     DWORD      ?
      szTip                     BYTE       128 dup(?)
      dwState                   DWORD      ?
      dwStateMask               DWORD      ?
      szInfo                    BYTE       256 dup(?)
      union
        uTimeout                DWORD      ?
        uVersion                DWORD      ?
      ends
      szInfoTitle               BYTE       64 dup(?)
      dwInfoFlags               DWORD      ?
    NOTIFYICONDATAAA ENDS
    
    NOTIFYICONDATA              equ  <NOTIFYICONDATAAA>
    ;Equipo\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced
    ;Set (or add) EnableBalloonTips (as REG_DWORD) and set value to 1 and Reeboot the PC

    ;NOTIFYICONDATA              equ  <NOTIFYICONDATAAA>
    ;note NOTIFYICONDATA <>
    ;
    ;icc INITCOMMONCONTROLSEX <INITCOMMONCONTROLSEX,ICC_BAR_CLASSES>
    icex                        INITCOMMONCONTROLSEX <> ;structure for Controls
    TrayBalloonNote             NOTIFYICONDATA <>
    ;
    ti                          TOOLINFO<>

    hToolTip                    dd ?
    ;
    szToolTipTxt_Delay_Solved   dd 2000 dup(?)
    ;
    pnt                         POINT <>
    hTrayPopupMenuRClick        dd ?
    hTrayPopupMenuLClick        dd ?
    bIconized                   db ?
    ;bHotKey                     db ?
    bHotKeyOn                   db ?
    ;#define                    CCSIZEOF_STRUCT(structname,member)
    ;#define                    TTTOOLINFOA_V1_SIZE CCSIZEOF_STRUCT(TTTOOLINFOA, lpszText)
    ;#define                    TTTOOLINFOW_V1_SIZE CCSIZEOF_STRUCT(TTTOOLINFOW, lpszText)
    ;#define                    TTTOOLINFOA_V2_SIZE CCSIZEOF_STRUCT(TTTOOLINFOA, lParam)
    ;#define                    TTTOOLINFOW_V2_SIZE CCSIZEOF_STRUCT(TTTOOLINFOW, lParam)
    ;#define                    TTTOOLINFOA_V3_SIZE CCSIZEOF_STRUCT(TTTOOLINFOA, lpReserved)
    ;#define                    TTTOOLINFOW_V3_SIZE CCSIZEOF_STRUCT(TTTOOLINFOW, lpReserved)
    ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    icce                       INITCOMMONCONTROLSEX<>
    ;
;    MOUSEINPUT      struct
;      ddx             dd ? ;dx
;      ddy             dd ? ;dy
;      mouseData       dd ?
;      dwFlags         dd ?
;      time            dd ?
;      dwExtraInfo     dd ?
;    MOUSEINPUT      ends
;    ;
;    KEYBDINPUT      struct
;      wVk             dw ?
;      wScan           dw ?
;      dwFlags         dd ?
;      time            dd ?
;      dwExtraInfo     dd ?
;    KEYBDINPUT      ends
;    ;
;    HARDWAREINPUT   struct
;      uMsg            dd ?
;      wParamL         dw ?
;      wParamH         dw ?
;    HARDWAREINPUT   ends
;    ;
    INPUT                       struct
      ddType                        dd ?  ;INPUT_MOUSE,INPUT_KEYBOARD,INPUT_HARDWARE
      UNION
        mi                          MOUSEINPUT <>
        ki                          KEYBDINPUT <>
        hi                          HARDWAREINPUT <>
      ENDS
    INPUT                       ends
    ;input                      INPUT <>
    ;
    BtnStruct                   XXBUTTON <?>
    ;hBmp                        dd ?
    hBmpNoiseNormal             dd ?
    hBmpNoiseHover              dd ?
    DrawingRec                  RECT<>
    WinMainRec                  RECT<>
    UpdateRec                   RECT<>
    ;
    szActivityLogText           db 32768 dup(?) ;Before EM_SETLIMITTEXT is called, the limit for the amount of text a user can enter in an edit control is 32,767 character
.code
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;PROGRAM ENTRY POINT
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
start:
    ;
    Invoke UpdateRegistryPrintScreen,FALSE
    mov dPrintScreenKeyForSnipping,eax
    ;
    Invoke RunCmdAtMainEntry
    .if eax==NULL ;Result NULL if NOT instance running, else result is hWin of instance running.
        call WinMain
        Invoke GlobalFindAtom,addr szAtom
        .if eax
            Invoke GlobalDeleteAtom,eax
        .endif
    .endif
    ;
    .if dPrintScreenKeyForSnipping
        ;Restore Registry Setting if Print Screen for Snipping was disable by this application.
        Invoke UpdateRegistryPrintScreen,dPrintScreenKeyForSnipping
    .endif
    xor eax,eax
    Invoke ExitProcess,eax
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
RunCmdAtMainEntry               proc

    mov hInstance,          FUNC(GetModuleHandle,NULL)
    mov hIconTray,          FUNC(LoadIcon,hInstance,500)
    mov hIcon,              FUNC(LoadIcon,hInstance,500)
    mov hCursorRed,         FUNC(LoadCursor,hInstance,600)
    ;
    mov hWinMain,NULL
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke UpdateActivityLog,NULL,NULL,0,0 ;Clear LOG
    szText szActivity_ApplicationInitialization,9,9,9,"***Application Initialization***"
    Invoke UpdateActivityLog,addr szActivity_ApplicationInitialization,NULL,0,0
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    ;************ Get Cursor ******
    szText szCursorFileName,"cursor.cur"
    Invoke GetCL,0,addr szCurFile
    Invoke lstrlen, addr szCurFile
    lea ecx,szCurFile          ;Change Name .exe to .ini
    add eax,ecx
    @@:
        dec eax                       ;Decrement the string
        cmp byte ptr [eax],'\'
        jne @B
    inc eax
    mov byte ptr [eax],0
    Invoke lstrcat,addr szCurFile,addr szCursorFileName
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    szText szActivity_StartCursor,"-Loading Cursor file: "
    Invoke UpdateActivityLog,addr szActivity_StartCursor,addr szCurFile,1,0
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke LoadImage,0,addr szCurFile,IMAGE_CURSOR,0,0,LR_LOADFROMFILE
    mov hCursor,eax
    .if hCursor
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        szText szActivity_StartCursorSucess,"-Load Cursor Sucess."
        Invoke UpdateActivityLog,addr szActivity_StartCursorSucess,NULL,1,0
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    .else
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetLastError
        Invoke aaerrortext,eax
        szText szActivity_StartCursorFail,"-Load Cursor fail with error: "
        Invoke UpdateActivityLog,addr szActivity_StartCursorFail,eax,1,0
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        ;
        Invoke GetCL,0,addr szCurFile
        Invoke lstrlen, addr szCurFile
        lea ecx,szCurFile          ;Change Name .exe to .cur
        mov byte ptr [ecx+eax-1],'r'
        mov byte ptr [ecx+eax-2],'u'
        mov byte ptr [ecx+eax-3],'c'
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        szText szActivity_StartAltCursor,"-Loading Alternative Cursor file: "
        Invoke UpdateActivityLog,addr szActivity_StartCursor,addr szCurFile,1,0
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke LoadImage,0,addr szCurFile,IMAGE_CURSOR,0,0,LR_LOADFROMFILE
        mov hCursor,eax
        .if hCursor
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            ;szText szActivity_StartCursorSucess,"-Load Cursor Sucess."
            Invoke UpdateActivityLog,addr szActivity_StartCursorSucess,NULL,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .else
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke GetLastError
            Invoke aaerrortext,eax
            ;szText szActivity_StartCursorFail,"->Load Cursor fail with error:"
            Invoke UpdateActivityLog,addr szActivity_StartCursorFail,eax,1,0
            ;
            szText szActivity_DefaultCursorFailTitle,"-Operation Info: "
            szText szActivity_DefaultCursorFailText,"Added internal 'Red Arrow' option for 'Capture The Mouse Cursor'"
            Invoke UpdateActivityLog,addr szActivity_DefaultCursorFailTitle,addr szActivity_DefaultCursorFailText,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .endif
    .endif
    ;
    ;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Invoke IniManager,0,0 ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)

    ;
    ;Invoke InitCommonControls
;        ICC_FLAGS = ICC_WIN95_CLASSES
;        ;comment out the ones you do not want
;        ICC_FLAGS = ICC_FLAGS or ICC_ANIMATE_CLASS
;        ICC_FLAGS = ICC_FLAGS or ICC_BAR_CLASSES
;        ICC_FLAGS = ICC_FLAGS or ICC_COOL_CLASSES
;        ICC_FLAGS = ICC_FLAGS or ICC_DATE_CLASSES
;        ICC_FLAGS = ICC_FLAGS or ICC_HOTKEY_CLASS
;        ICC_FLAGS = ICC_FLAGS or ICC_INTERNET_CLASSES
;        ICC_FLAGS = ICC_FLAGS or ICC_LISTVIEW_CLASSES
;        ICC_FLAGS = ICC_FLAGS or ICC_PAGESCROLLER_CLASS
;        ICC_FLAGS = ICC_FLAGS or ICC_PROGRESS_CLASS
;        ICC_FLAGS = ICC_FLAGS or ICC_TAB_CLASSES
;        ICC_FLAGS = ICC_FLAGS or ICC_TREEVIEW_CLASSES
;        ICC_FLAGS = ICC_FLAGS or ICC_UPDOWN_CLASS
;        ICC_FLAGS = ICC_FLAGS or ICC_USEREX_CLASSES
    mov icce.dwSize, SIZEOF INITCOMMONCONTROLSEX
    mov icce.dwICC,ICC_HOTKEY_CLASS+ICC_UPDOWN_CLASS
    invoke InitCommonControlsEx,addr icce
    ;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Invoke GetCL,0,addr szProgramPath
    ;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ;
    jmp @F
        mov bWinMainInstanceHotSendKeys,TRUE
        mov hWinMainInstanceAtom,NULL
        mov hPreviousWin,NULL
        Invoke FindApplicationProcessID,addr szProgramPath,1,addr hPreviousWin     ;dWhatToDoIfFound: ReturnWinHandle=-1,Hide=0, Unhide=1, Kill=2
        .if hWinMainInstanceAtom
            mov eax,hWinMainInstanceAtom
            jmp ExitRunCmdAtMainEntry
        .else
            jmp OmmitAtomCheck
        .endif
    @@:

    mov hWinMainInstanceAtom,NULL;
    mov eax,TRUE ;TRUE=DEBUG to clear Atom
    .if eax==TRUE
        jmp @F
        Invoke GlobalFindAtom,addr szAtom
        Invoke GlobalDeleteAtom,eax
        jmp OmmitAtomCheck
        @@:
    .endif
    ;Check if CTRL KEY is Pressed
    Invoke GetKeyState,VK_CONTROL ;VK_MENU ;VK_SHIFT ;VK_CONTROL ;VK_CONTROL = RIGHT_CTRL_PRESSED || LEFT_CTRL_PRESSED
    .if (ax > 1) ;If the high-order bit is 1, the key is down; otherwise, it is up.
        .if eax
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            szText szActivity_StartCTRKEY_Title,"-Debug key:"
            szText szActivity_StartCTRKEY_Text," CTRL key pressed = TRUE"
            Invoke UpdateActivityLog,addr szActivity_StartCTRKEY_Title,addr szActivity_StartCTRKEY_Text,2,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke GlobalFindAtom,addr szAtom
            .if eax  ;If the function succeeds, the return value is the global atom associated with the given string, If the function fails, the return value is zero.
                Invoke GlobalDeleteAtom,eax
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                szText szActivity_StartCTRKEY_Atom,"-Find Atom:"
                ;Invoke UpdateActivityLog,addr szAtom,NULL,2
                szText szActivity_StartCTRKEY_AtomFound," Atom was found and Deleted."
                Invoke UpdateActivityLog,addr szActivity_StartCTRKEY_Atom,addr szActivity_StartCTRKEY_AtomFound,1,0
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            .else
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                ;szText szActivity_StartCTRKEY_Atom,"-Find Atom:"
                ;Invoke UpdateActivityLog,addr szAtom,NULL,2
                szText szActivity_StartCTRKEY_AtomNotFound," Atom was not found"
                Invoke UpdateActivityLog,addr szActivity_StartCTRKEY_Atom,addr szActivity_StartCTRKEY_AtomNotFound,1,0
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            .endif
        .endif
    .endif
    ;
    Invoke GlobalFindAtom,addr szAtom ;If the function fails, the return value is zero, If the function succeeds, the return value is the atom associated with the given string.
    .if eax == 0 ;If the function fails, the return value is zero
        @@:
        ;Invoke ErrorManager, addr szAppTitle
        Invoke GlobalAddAtom,addr szAtom
        mov hSingleInstanceAtom,eax
        .if eax!=NULL
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            szText szActivity_StartPreviousInstanceEntryPoint,"-Entry point: "
            szText szActivity_StartPreviousInstanceNewAtom,"an unique instance Id was created."
            Invoke UpdateActivityLog,addr szActivity_StartPreviousInstanceEntryPoint,addr szActivity_StartPreviousInstanceNewAtom,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .else
            Invoke GetLastError
            Invoke aaerrortext,eax
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            szText szActivity_StartPreviousInstanceNoAtom,"-Failed to create Unique instance Id:"
            Invoke UpdateActivityLog,addr szActivity_StartPreviousInstanceEntryPoint,addr szActivity_StartPreviousInstanceNoAtom,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .endif
    .else
        mov hSingleInstanceAtom,eax
        mov eax,hSingleInstanceAtom
        .if eax!=0
            ;
            mov bWinMainInstanceHotSendKeys,TRUE
            mov hPreviousWin,NULL
            Invoke FindApplicationProcessID,addr szProgramPath,1,addr hPreviousWin     ;dWhatToDoIfFound: ReturnWinHandle=-1,Hide=0, Unhide=1, Kill=2
            ;
;            .if hWinMainInstanceAtom != NULL
;                mov hWinMainInstanceAtom,eax
;                jmp ExitRunCmdAtMainEntry
;                Invoke IsIconic,hWinMainInstanceAtom
;                .if eax
;                    Invoke ShowWindow,hWinMainInstanceAtom,SW_RESTORE
;                .endif
;                Invoke IsWindowVisible,hWinMainInstanceAtom
;                .if eax
;                    Invoke ShowWindow,hWinMainInstanceAtom,SW_SHOW
;                    Invoke SetForegroundWindow,hWinMainInstanceAtom
;                .endif
;                Invoke SetWindowPos,hWinMainInstanceAtom,HWND_TOP,0,0,0,0,SWP_SHOWWINDOW OR SWP_NOREPOSITION OR SWP_NOSENDCHANGING OR SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOZORDER
;;;                ;Invoke MessageBox,NULL,addr szProgramPath,addr szProgramPath,NULL
;;;                Invoke ShowWindowAsync,hWinMainInstanceAtom, SW_SHOWNORMAL or SW_RESTORE
;;;                Invoke SetForegroundWindow,hWinMainInstanceAtom
;           .endif
            xor eax,eax
            inc eax
            jmp ExitRunCmdAtMainEntry
            Invoke ExitProcess,eax
        .endif
    .endif
    ;
    OmmitAtomCheck:
    ;mov icex.dwSize,sizeof INITCOMMONCONTROLSEX
    ;mov icex.dwICC,ICC_WIN95_CLASSES
    ;Invoke InitCommonControlsEx,addr icex  ;Windows XP: If a manifest is used, InitCommonControlsEx is not required.
    ;Invoke RegisterClassEx,addr wc
    ;
    ;mov icex.dwSize,sizeof icex
    ;mov icex.dwICC,ICC_BAR_CLASSES+ICC_STANDARD_CLASSES+ICC_TAB_CLASSES+ICC_WIN95_CLASSES
    ;invoke  InitCommonControlsEx,addr icex
    ;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ;Invoke GetCL,0,addr szProgramPath
    ;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ;************ Get Full Path For Ini Archiv ******
    Invoke GetCL,0,addr szIniFile
    Invoke lstrlen, addr szIniFile
    lea ecx,szIniFile          ;Change Name .exe to .ini
    mov byte ptr [ecx+eax-1],'i'
    mov byte ptr [ecx+eax-2],'n'
    mov byte ptr [ecx+eax-3],'i'
    Invoke CreateFile,addr szIniFile,GENERIC_READ,FILE_SHARE_READ,0,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,0
    .if eax == INVALID_HANDLE_VALUE
        xor eax,eax
    .else
        Invoke CloseHandle,eax
        xor eax,eax
        add eax,1
    .endif
    mov bIniArchiveExists,al

    ;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ;************ Get Full Quote Path For Register at Run ******
    Invoke lstrcpy,addr szQuotedPath,addr szQuote
    Invoke lstrcat,addr szQuotedPath,addr szProgramPath
    Invoke lstrcat,addr szQuotedPath,addr szQuote
    ;Invoke MessageBoxA,NULL,addr szQuotedPath,addr szQuote,NULL
    ;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ;********** Get Only the ProgramName.exe
    Invoke lstrlen, addr szProgramPath
    lea ecx,szProgramPath
    add eax,ecx
    @@:
    dec eax                       ;Decrement the string
    cmp byte ptr [eax],'\'
    jne @B
    inc eax
    Invoke lstrcpy,addr szProgramName,eax
    ;
    Invoke lstrlen, addr szProgramName
    lea ecx,szProgramName
    add eax,ecx
    sub eax,4
    mov byte ptr [eax],00H                            ;Remove .exe
    ;
    ;szText szServiceName,"MyService"
    ;Invoke ServiceManager,addr szServiceName,0,0,REMOVE      ;PROTO :DWORD,:DWORD,:DWORD,:BYTE ;(ServiceName),(ServiceDesc),(ServicePath),(INSTALL,REMOVE,ISINSTALLED,ISRUNNING)
    ;
    xor eax, eax
    ExitRunCmdAtMainEntry:
    ret
RunCmdAtMainEntry               endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
;                       WINMAIN FOR SnapShot
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
WinMain                         proc
    ;WS_OVERLAPPED OR WS_POPUP OR WS_VISIBLE OR WS_CAPTION OR WS_SYSMENU OR DS_CENTER OR WS_EX_TOPMOST
    LOCAL lpArgs:DWORD
    Invoke GlobalAlloc,GMEM_FIXED or GMEM_ZEROINIT, 32
    mov lpArgs, eax
    push hIcon
    pop [eax]
            ; caption,font,pointsize MS Sans Serif  ;+ WS_EX_TOPMOST OR WS_VISIBLE  OR WS_EX_TOPMOST  OR WS_EX_APPWINDOW WS_OVERLAPPEDWINDOW WS_OVERLAPPED
            ;WS_OVERLAPPED OR WS_POPUP OR WS_VISIBLE OR WS_CAPTION OR WS_SYSMENU OR DS_CENTER OR WS_EX_TOPMOST
            ;WS_VISIBLE+
    Dialog  "CapturA+: Screen Capture & SnapShot Recorder","Segoe UI",9,\
            WS_OVERLAPPED+WS_POPUP+WS_CAPTION+WS_SYSMENU+WS_MINIMIZEBOX+DS_CENTER+WS_EX_TOPMOST,\
            96, \                                            ; control count
            0,0,540,298, \                                   ; x y co-ordinates
            8192                                             ; memory buffer size
            ;WS_OVERLAPPED WS_OVERLAPPEDWINDOW
            ;DlgIcon   500,260,62,101
            ;
            ;===============================================
            DlgButton "Group", BS_GROUPBOX OR WS_GROUP OR BS_FLAT OR CCS_NODIVIDER OR BS_CENTER OR SS_ETCHEDVERT,000,000,204,298,IDC_LIST_GROUP_OPTIONS
            DlgRadio "Capture",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP  OR BS_PUSHLIKE OR BS_LEFT                  ,000,000,204,032,IDC_LIST_ITEM_BUTTON1
            DlgRadio "File",   WS_CHILD OR WS_VISIBLE OR WS_TABSTOP  OR BS_PUSHLIKE OR BS_LEFT                  ,000,032,204,032,IDC_LIST_ITEM_BUTTON2 ;OR BS_LEFTTEXT
            DlgRadio "Application",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP  OR BS_PUSHLIKE OR BS_LEFT              ,000,064,204,032,IDC_LIST_ITEM_BUTTON3
            DlgRadio "Hotkeys",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP  OR BS_PUSHLIKE OR BS_LEFT                  ,000,096,204,032,IDC_LIST_ITEM_BUTTON4
            DlgRadio "Activity Log",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP  OR BS_PUSHLIKE OR BS_LEFT             ,000,128,204,032,IDC_LIST_ITEM_BUTTON5
            DlgRadio "About",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP  OR BS_PUSHLIKE OR BS_LEFT                    ,000,160,204,032,IDC_LIST_ITEM_BUTTON6
            DlgRadio "Save Settings",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_PUSHLIKE OR BS_LEFT             ,000,192,204,032,IDC_LIST_ITEM_BUTTON7
            ;DlgRadio "Quit application",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_PUSHLIKE OR BS_LEFT         ,000,224,204,032,IDC_GROUP_LIST_BUTTON8
            ;DlgRadio "Quit application",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_PUSHLIKE OR BS_LEFT         ,000,256,204,032,IDC_GROUP_LIST_BUTTON9
            DlgRadio "Quit application",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP  OR BS_PUSHLIKE OR BS_LEFT         ,000,254,204,36,IDC_LIST_ITEM_BUTTON8
            ;
            ;===============================================
            ;          propertySheet
            DlgComCtl "SysTabControl32",WS_CHILD OR WS_VISIBLE,204,0,336,298,IDC_SHEET_TAB_CONTROL
            ;===============================================
            ;
            ;===============================================
            ;           CAPTURE SETTINGS                    ;Delay x seconds. At the end of delay, click in a windows to snap it.
            ;===============================================
            DlgStatic "Capture",WS_CHILD OR WS_VISIBLE,204+12,12,312,20,IDC_PROMPT_TITLE_CAPTURE
            DlgStatic "Select How To Capture Screenshots",WS_CHILD OR WS_VISIBLE,204+12,40,312,16,IDC_PROMPT_SUBTITLE_CAPTURE
            ;
            DlgStatic "Select Capture Target",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,62,312,10,IDC_PROMPT_CAPTURE_TARGET
            DlgCombo CBS_DROPDOWNLIST OR CBS_AUTOHSCROLL OR CBS_HASSTRINGS OR CBS_NOINTEGRALHEIGHT OR WS_VSCROLL OR WS_TABSTOP,204+12,74,180,80,IDC_DROP_CAPTURE_TARGET
            DlgButton "Region",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP,204+12+180+4,74,36,12,IDC_BTN_DIALOG_CAPTURE_REGION
            ;
            DlgStatic "Capture The Mouse Cursor:",WS_CHILD OR WS_VISIBLE,204+12,92,312,10,IDC_PROMPT_CAPTURE_CURSOR
            DlgCombo CBS_DROPDOWNLIST OR CBS_AUTOHSCROLL OR CBS_HASSTRINGS OR CBS_NOINTEGRALHEIGHT OR WS_VSCROLL OR WS_TABSTOP,204+12,104,180,80,IDC_DROP_CAPTURE_CURSOR
            ;CopyCursor=CopyIcon (si es animado = CopyImage),WindowFromPoint
            ;
            DlgStatic "Image Mode",WS_CHILD OR WS_VISIBLE,204+12,122,312,10,IDC_PROMPT_COLOR_IMAGE_MODE
            DlgCombo CBS_DROPDOWNLIST OR CBS_AUTOHSCROLL OR CBS_HASSTRINGS OR CBS_NOINTEGRALHEIGHT OR WS_VSCROLL OR WS_TABSTOP,204+12,134,180,80,IDC_DROP_COLOR_IMAGE_MODE
            ;
            DlgStatic "Capture Mode",WS_CHILD OR WS_VISIBLE,204+12,154,312,10,IDC_PROMPT_CAPTURE_MODE
            DlgCombo CBS_DROPDOWNLIST OR CBS_AUTOHSCROLL OR CBS_HASSTRINGS OR CBS_NOINTEGRALHEIGHT OR WS_VSCROLL OR WS_TABSTOP,204+12,166,180,80,IDC_DROP_CAPTURE_MODE
            ;
            DlgStatic "Time Lapse Between Captures",WS_CHILD OR WS_VISIBLE,204+12,186,180,10,IDC_PROMPT_TIMELAPSE
            ;DlgDateTime  WS_TABSTOP or DTS_TIMEFORMAT,76,98,60,12,IDC_EDITBOX_TIMELAPSE_MIN
            ;UDS_WRAP           ;range 1 to 10: 123...8 9 10 begins with 1, 2 etc
            ;UDS_AUTOBUDDY      ;Auto fits into previous control.
            ;UDS_SETBUDDYINT    ;Works with Intergers
            ;UDS_ARROWKEYS      ;Looks like arrows
            ;UDS_ALIGNRIGHT     ;Arrows alinged right
            ;UDS_NOTHOUSANDS
            DlgEdit  WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_NUMBER OR ES_CENTER OR ES_AUTOHSCROLL,204+12,198,86,12,IDC_EDITBOX_TIMELAPSE_MIN
            DlgComCtl "msctls_updown32",UDS_AUTOBUDDY OR UDS_SETBUDDYINT OR UDS_ARROWKEYS OR UDS_ALIGNRIGHT OR UDS_NOTHOUSANDS,204+12,198,12,12,IDC_SPINBOX_TIMELAPSE_MIN
            DlgEdit  WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_NUMBER OR ES_CENTER OR ES_AUTOHSCROLL,204+12+94,198,86,12,IDC_EDITBOX_TIMELAPSE_SEC
            DlgComCtl "msctls_updown32",UDS_AUTOBUDDY OR UDS_SETBUDDYINT OR UDS_ARROWKEYS OR UDS_ALIGNRIGHT OR UDS_NOTHOUSANDS,204+12+94,198,12,12,IDC_SPINBOX_TIMELAPSE_SEC
            ;
            DlgStatic "Start 'Auto Shot' Capture In Application Init",WS_CHILD OR WS_VISIBLE,204+12,218,312,10,IDC_PROMPT_START_AUTOSHOT_AT_INIT
            DlgCheck "Off",WS_CHILD OR WS_TABSTOP OR BS_AUTOCHECKBOX,204+12,230,54,12,IDC_CHK_START_AUTOSHOT_AT_INIT
            ;
            DlgStatic "Click Snap Delay",WS_CHILD OR WS_VISIBLE,204+12+180+8,154,134,10,IDC_PROMPT_CLICK_DELAY
            DlgEdit  WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_NUMBER OR ES_CENTER OR ES_AUTOHSCROLL,204+12+180+8,166,86,12,IDC_EDITBOX_CLICK_DELAY
            DlgComCtl "msctls_updown32", UDS_AUTOBUDDY OR UDS_SETBUDDYINT OR UDS_ARROWKEYS OR UDS_ALIGNRIGHT OR UDS_NOTHOUSANDS,204+12+180+8,166,12,12,IDC_SPINBOX_CLICK_DELAY
            ;
            DlgButton "Start Auto Shot",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_AUTOCHECKBOX,204+12+000,254,060,36,IDC_BTN_AUTOSHOT
            DlgButton "Wait 5 seconds, then open Location",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP,   204+12+062,254,060,36,IDC_BTN_SEND_TO_LOCATION
            DlgButton "Wait 5 seconds, then new Email",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP,      204+12+124,254,060,36,IDC_BTN_SEND_TO_EMAIL
            DlgButton "Wait 5 seconds, then Clipboard",WS_CHILD OR WS_TABSTOP,                     204+12+186,254,060,36,IDC_BTN_SEND_TO_CLIPBOARD
            DlgButton "Close",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_DEFPUSHBUTTON,         204+12+248,254,060,36,IDC_BTN_HIDE
            ;DlgStatic WS_CHILD OR WS_VISIBLE OR WS_POPUP and WS_EX_TOOLWINDOW,10,10,10,10,IDC_TOOLTIP ;WS_POPUP OR TTS_NOPREFIX OR TTS_BALLOON
            ;
            ;DlgButton "HotKey...",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP  OR 8000H OR WS_DISABLED                 ,016,222,056,12,IDC_BTN_HOTKEY
            ;DlgButton "Region...",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP  OR 8000H                                ,076,222,060,12,IDC_BTN_SEND_TO_EMAIL
            ;DlgStatic "Shot HotKey:",WS_CHILD OR WS_VISIBLE,16,115,58,12,IDC_PROMPT_IMAGE_ZOOM
            ;DlgButton "Default",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP  OR 8000H,16,130,60,12,IDC_BTN_DIALOG_ARCHIVE_PATH
            ;DlgComCtl "SysDateTimePick32",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP,76,114,60,12,HOTKEY_PRINTSCREEN
            ;
            ;DlgStatic "App. HotKey:",WS_CHILD OR WS_VISIBLE,16,131,58,12,IDC_PROMPT_TARGET_FILE_FORMAT3
            ;DlgButton "Configure HotKey",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP  OR 8000H,76,130,60,12,IDC_BTN_HOTKEY
            ;DlgComCtl "msctls_hotkey32",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP,76,130,60,12,IDC_HOTKEY2
            ;
            ;===============================================
            ;           FILE SETTINGS
            ;===============================================
            DlgStatic "File",WS_CHILD OR WS_VISIBLE,204+12,12,312,20,IDC_PROMPT_TITLE_FILE
            DlgStatic "Copy File Path To Clipboard After Capture",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,40,312,10,IDC_PROMPT_COPY_PATH_TO_CLIPBOARD
            DlgButton "Off",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_PUSHLIKE OR BS_AUTOCHECKBOX,204+12,52,54,12,IDC_CHK_COPY_PATH_TO_CLIPBOARD
            DlgStatic "Choose File Settings",WS_CHILD OR WS_VISIBLE,204+12,72,312,16,IDC_PROMPT_SUBTITLE_FILE
            DlgStatic "Select Screenshot File Format",WS_CHILD OR WS_VISIBLE,204+12,94,312,10,IDC_PROMPT_TARGET_FILE_FORMAT
            DlgCombo CBS_DROPDOWNLIST OR CBS_AUTOHSCROLL OR CBS_HASSTRINGS OR CBS_NOINTEGRALHEIGHT OR WS_VSCROLL OR WS_TABSTOP,204+12,106,180,40,IDC_DROP_TARGET_FILE_FORMAT
            DlgStatic "File Path",WS_CHILD OR WS_VISIBLE,204+12,126,312,10,IDC_PROMPT_ARCHIVE_PATH
            DlgEdit  WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL OR 8000H,204+12,138,272,12,IDC_EDITBOX_ARCHIVE_PATH
            DlgButton "Browse",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP,204+12+276,138,36,12,IDC_BTN_DIALOG_ARCHIVE_PATH
            DlgStatic "Shot Sequence (%Counter%)",WS_CHILD OR WS_VISIBLE,204+12,158,312,10,IDC_PROMPT_COUNTER
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_NUMBER OR ES_CENTER OR ES_AUTOHSCROLL,204+12,170,180,12,IDC_EDITBOX_COUNTER
            DlgComCtl "msctls_updown32",UDS_AUTOBUDDY OR UDS_SETBUDDYINT OR UDS_ARROWKEYS OR UDS_ALIGNRIGHT OR UDS_NOTHOUSANDS,204+12+180,170,12,12,IDC_SPINBOX_COUNTER
            DlgStatic "JPG Image Quality (File Size)",WS_CHILD OR WS_VISIBLE,204+12,190,312,10,IDC_PROMPT_IMAGE_QUALITY
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR ES_NUMBER OR ES_CENTER OR WS_TABSTOP OR ES_AUTOHSCROLL,204+12,202,180,12,IDC_EDITBOX_IMAGE_QUALITY
            DlgComCtl "msctls_updown32",UDS_AUTOBUDDY OR UDS_SETBUDDYINT OR UDS_ARROWKEYS OR UDS_ALIGNRIGHT OR UDS_NOTHOUSANDS,204+12+180,204,12,12,IDC_SPINBOX_IMAGE_QUALITY
            ;
            DlgStatic "JPG Image Dimensions: Zoom In or Zoom Out (%)",WS_CHILD OR WS_VISIBLE,204+12,222,312,10,IDC_PROMPT_IMAGE_ZOOM
            DlgEdit  WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_NUMBER OR ES_CENTER OR ES_AUTOHSCROLL,204+12,234,180,12,IDC_EDITBOX_IMAGE_ZOOM
            DlgComCtl "msctls_updown32",UDS_AUTOBUDDY OR UDS_SETBUDDYINT OR UDS_ARROWKEYS OR UDS_ALIGNRIGHT OR UDS_NOTHOUSANDS,204+12+180,234,12,12,IDC_SPINBOX_IMAGE_ZOOM
            ;
            DlgStatic "JPG Comments",WS_CHILD OR WS_VISIBLE,204+12,254,312,10,IDC_PROMPT_JPG_COMMENTS
            DlgEdit  WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL OR 8000H,204+12,266,312,12,\
                      IDC_EDITBOX_JPG_COMMENTS
            ;===============================================
            ;           APPLICATION SETTINGS
            ;===============================================
            DlgStatic "Application",WS_CHILD OR WS_VISIBLE,204+12,12,312,20,                                            IDC_PROMPT_TITLE_APP
            DlgStatic "Choose Application Settings",WS_CHILD OR WS_VISIBLE,204+12,40,312,16,                            IDC_PROMPT_SUBTITLE_APP
            DlgStatic "Run When Windows Starts",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,62,312,10,                    IDC_PROMPT_WIN_STARTUP_RUN
            DlgButton "Off",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_PUSHLIKE OR BS_AUTOCHECKBOX,204+12,74,54,12,     IDC_CHK_WIN_STARTUP_RUN
            DlgStatic "Hide Application Window At Opening Or Minimizing",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,94,312,10,IDC_PROMPT_HIDE_WIN_WHEN_OPENING
            DlgButton "Off",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_PUSHLIKE OR BS_AUTOCHECKBOX,204+12,106,54,12,    IDC_CHK_HIDE_WIN_WHEN_OPENING
            DlgStatic "Audible Notification After Capture",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,126,312,10,        IDC_PROMPT_PLAY_AUDIBLE_SHOT
            DlgButton "Off",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_PUSHLIKE OR BS_AUTOCHECKBOX,204+12,138,54,12,    IDC_CHK_PLAY_AUDIBLE_SHOT
            DlgStatic "Open Last Capture Location Path",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,158,312,10,           IDC_PROMPT_OPEN_LAST_CAPTURE_PATH
            DlgButton "Off",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_PUSHLIKE OR BS_AUTOCHECKBOX,204+12,170,54,12,    IDC_CHK_OPEN_LAST_CAPTURE_PATH
            DlgStatic "Copy Screenshot Image To Clipboard",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,190,312,10,        IDC_PROMPT_COPY_SCREENSHOT_CLIPBOARD
            DlgButton "Off",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_PUSHLIKE OR BS_AUTOCHECKBOX,204+12,202,54,12,    IDC_CHK_COPY_SCREENSHOT_CLIPBOARD
            DlgStatic "Show A Notification Icon In System Tray",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,222,312,10,   IDC_PROMPT_SHOW_NOTIFICATION_ICON
            DlgButton "Off",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_PUSHLIKE OR BS_AUTOCHECKBOX,204+12,234,54,12,    IDC_CHK_SHOW_NOTIFICATION_ICON
            DlgStatic "Display A Notification Balloon After Capture",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,254,312,10,      IDC_PROMPT_DISPLAY_NOTIFICATION_CAPTURE
            DlgButton "Off",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR BS_PUSHLIKE OR BS_AUTOCHECKBOX,204+12,266,54,12,    IDC_CHK_DISPLAY_NOTIFICATION_CAPTURE
            ;https://learn.microsoft.com/en-us/windows/win32/shell/notification-area#display-the-notification
            ;Minimize Window Into Notification Icon
            ;
            ;===============================================
            ;           HOTKEYS
            ;===============================================
            DlgStatic "Hotkeys",WS_CHILD OR WS_VISIBLE,204+12,12,312,20,                                                    IDC_PROMPT_HOTKEYS_TITLE
            DlgStatic "Active Registered Hotkeys",WS_CHILD OR WS_VISIBLE,204+12,40,312,16,                                  IDC_PROMPT_HOTKEYS_SUBTITLE
            DlgStatic "Unhide This Application",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,62,312,10,                        IDC_PROMPT_HOTKEYS_SHOW_ME_LEFT
            DlgEdit  ES_READONLY OR WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL,204+12,74,180,12,   IDC_EDITBOX_HOTKEYS_SHOW_ME_LEFT
            DlgStatic "Capture Target",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,94,312,10,                                 IDC_PROMPT_HOTKEYS_CAPTURE_TARGET
            DlgEdit  ES_READONLY OR WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL,204+12,106,180,12,  IDC_EDITBOX_HOTKEYS_CAPTURE_TARGET
            DlgStatic "Capture Target And Open Default Editor",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,126,312,10,        IDC_PROMPT_HOTKEYS_FOR_EDIT_TARGET
            DlgEdit  ES_READONLY OR WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL,204+12,138,180,12,  IDC_EDITBOX_HOTKEYS_FOR_EDIT_TARGET
            DlgStatic "Capture Active Window",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,158,312,10,                         IDC_PROMPT_HOTKEYS_CAPTURE_TOPWIN
            DlgEdit  ES_READONLY OR WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL,204+12,170,180,12,  IDC_EDITBOX_HOTKEYS_CAPTURE_TOPWIN
            DlgStatic "Capture Active Window And Open Default Editor",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,190,312,10, IDC_PROMPT_HOTKEYS_FOR_EDIT_TOPWIN
            DlgEdit  ES_READONLY OR WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL,204+12,202,180,12,  IDC_EDITBOX_HOTKEYS_FOR_EDIT_TOPWIN
            DlgStatic "Capture Target And Open Email Client",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,222,312,10,          IDC_PROMPT_HOTKEYS_EMAIL_TARGET
            DlgEdit  ES_READONLY OR WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL,204+12,234,180,12,  IDC_EDITBOX_HOTKEYS_EMAIL_TARGET
            DlgStatic "Capture Active Window And Open Email Client",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,254,312,10,   IDC_PROMPT_HOTKEYS_EMAIL_TOPWIN
            DlgEdit  ES_READONLY OR WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOVSCROLL OR WS_VSCROLL OR ES_MULTILINE,204+12,266,180,12,  IDC_EDITBOX_HOTKEYS_EMAIL_TOPWIN
            ;ES_AUTOHSCROLL=SingleLine
            ;===============================================
            ;           ACTIVITY LOG
            ;===============================================
            DlgStatic "Activity Log",WS_CHILD OR WS_VISIBLE,204+12,12,312,20,                                               IDC_PROMPT_ACTIVITY_TITLE
            DlgStatic "View Monitoring And Troubleshooting Messsages",WS_CHILD OR WS_VISIBLE,204+12,40,312,16,              IDC_PROMPT_ACTIVITY_SUBTITLE
            DlgStatic "Last Activity Info",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,62,312,10,                             IDC_PROMPT_ACTIVITY_LOG
            DlgEdit  ES_READONLY OR WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL OR WS_VSCROLL OR ES_MULTILINE,204+12,74,312,204,  IDC_EDITBOX_ACTIVITY_LOG
            ;===============================================
            ;           ABOUT
            ;===============================================
            DlgStatic "About",WS_CHILD OR WS_VISIBLE,204+12,12,312,20,                                                      IDC_PROMPT_ABOUT_TITLE
            DlgStatic "Screen Capture Tool",WS_CHILD OR WS_VISIBLE,204+12,40,312,16,                                        IDC_PROMPT_ABOUT_SUBTITLE
            DlgStatic "Application History",WS_GROUP OR WS_CHILD OR WS_VISIBLE,204+12,62,312,10,                            IDC_PROMPT_ABOUT_LOG
            DlgEdit  ES_READONLY OR WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOVSCROLL OR WS_VSCROLL OR ES_MULTILINE,204+12,74,312,204,  IDC_EDITBOX_ABOUT_LOG
            ;===============================================
            ;          propertySheet
            ;DlgComCtl "SysTabControl32",WS_CLIPSIBLINGS OR WS_CHILD OR WS_VISIBLE,204,0,336,298,IDC_SHEET_TAB_CONTROL_COVER
            ;===============================================
            ;
        CallModalDialog hInstance,0,WinMainProc,addr lpArgs
        pop esi
        Invoke GlobalFree,lpArgs
        ;
    ret
WinMain                         endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
;                       DIALOG PROCEDURE FOR SnapShot
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
WinMainProc                     proc hWnd:DWORD,uMsg:DWORD,wParam:DWORD,lParam:DWORD
                                ;
    .if uMsg == WM_INITDIALOG ;WM_CREATE
        ;
        ;
        Invoke SetWindowLong,hWnd,GWL_USERDATA,lParam
        mov eax, lParam
        mov eax, [eax]
        Invoke SendMessage,hWnd,WM_SETICON,1,[eax]
        ;
        Invoke InitDialogWinMain,hWnd
        ;
        Invoke GetThemeFonts
        Invoke SetTheme,hWnd,NULL
        Invoke SelectPage,IDC_PAGE01
        ;
        .if bCaptureMode==RUNTIME_CAPTURE_MODE_AUTOSHOT && bAutoShotStartAtInit==TRUE
            Invoke GetDlgItem,hWnd,IDC_BTN_AUTOSHOT
            push eax
            Invoke SendMessage,eax,BM_SETCHECK,BST_CHECKED,NULL
            ;Invoke SendMessage,eax,BM_SETCHECK,BST_CHECKED,NULL ;TODO: Cambiar por un pushed?
            pop eax
            Invoke SetWindowText,eax,addr szText_StopAutoShot
            Invoke StartAutoShotLoop,NULL,bAutoShot_TimerMinutes,bAutoShot_TimerSeconds,FALSE
        .endif
        ;
        ;.if uMsg == WM_INITDIALOG Return value
        ;The dialog box procedure should return TRUE to direct the system to set the keyboard focus to the control specified by wParam.
        mov eax,TRUE
        .if bHideApplicationWindowWhenOpening && (bIsUnhideApplicationWindowHotKeyActive  || bShowANotificationIcon)
            mov dIsWindowStyleTOOLWINDOW,FALSE
            .if bShowANotificationIcon
                mov dFirstTimeShowed,TRUE
                Invoke ShowWindow,hWinMain,SW_MINIMIZE
                Invoke GetWindowLong,hWnd,GWL_EXSTYLE
                mov dIsWindowStyleTOOLWINDOW,eax
                and eax,WS_EX_APPWINDOW
                or eax,WS_EX_TOOLWINDOW
                Invoke SetWindowLong,hWinMain,GWL_EXSTYLE,eax
                @@:
            .else
                mov dFirstTimeShowed,FALSE
                Invoke ShowWindow,hWinMain,SW_HIDE
                Invoke ShowWindow,hWinMain,SW_MINIMIZE ;SW_SHOWMINNOACTIVE ;SW_HIDE
            .endif                
            xor eax,eax
            ;Otherwise, it should return FALSE to prevent the system from setting the default keyboard focus.
        .endif
        ret
    ;.elseif uMsg == WM_SETFOCUS ;wParam A handle to the window that has lost the keyboard focus. This parameter can be NULL
;    .elseif uMsg == WM_ACTIVATE ;wParam The low-order word specifies whether the window is being activated or deactivated.
;                                ;lParam A handle to the window being activated or deactivated, depending on the value of the wParam parameter.
;                                ;WA_INACTIVE: Deactivated, WA_CLICKACTIVE: Activated by a mouse click, WA_ACTIVE: Activated by some method other than a mouse click
;        mov eax,wParam
;        .if al != WA_INACTIVE ;Activated by some method other than a mouse click (for example, by a call to the SetActiveWindow function or by use of the keyboard interface to select the window)
;            ;Invoke GetCurrentProcess                  ;Free Unuse Memory
;            ;Invoke SetProcessWorkingSetSize,eax,-1,-1 ;Free Unuse Memory
;            .if bHideApplicationWindowWhenOpening && bIsUnhideApplicationWindowHotKeyActive && bShowANotificationIcon==FALSE
;                Invoke MessageBeep,NULL
;                ;Windows was hidden at  InitDialogWinMain, now hide the Application ICON at Taskbar
;                ;Invoke ShowWindow,hWinMain,SW_MINIMIZED
;                ;Invoke SendMessage,hWinMain,WM_SYSCOMMAND,SC_MINIMIZE,0
;                Invoke ShowWindowAsync,hWinMain,SW_HIDE
;            .endif
;        .endif
;    .elseif uMsg == WM_ERASEBKGND
;        invoke GetClientRect, hWnd, addr Rct
;        mov hBrush, $invoke (CreatePatternBrush,hBmpNoiseNormal)
;        invoke SelectObject, wParam, hBrush
;        mov hOldBrush, eax
;        invoke PatBlt, wParam, Rct.left, Rct.top, Rct.right, Rct.bottom, PATCOPY
;        invoke SelectObject, wParam, hOldBrush
;        invoke DeleteObject, hBrush
;        return TRUE
;    .elseif uMsg==WM_WINDOWPOSCHANGED
;        Invoke GetWindowRect, hWnd,addr WinMainRec
;        Invoke RedrawWindow,0,addr WinMainRec,0,RDW_INVALIDATE OR RDW_ALLCHILDREN OR RDW_UPDATENOW
    .elseif uMsg == WM_CTLCOLORSTATIC
        ;.if bSelectingPage == FALSE
           Invoke GetDlgCtrlID, lParam
          .if eax == IDC_EDITBOX_HOTKEYS_SHOW_ME_LEFT || eax == IDC_EDITBOX_HOTKEYS_CAPTURE_TARGET || \
                eax == IDC_EDITBOX_HOTKEYS_FOR_EDIT_TARGET || eax == IDC_EDITBOX_HOTKEYS_CAPTURE_TOPWIN || \
                eax == IDC_EDITBOX_HOTKEYS_FOR_EDIT_TOPWIN || eax == IDC_EDITBOX_HOTKEYS_EMAIL_TARGET || \
                eax == IDC_EDITBOX_HOTKEYS_EMAIL_TOPWIN || eax == IDC_EDITBOX_ACTIVITY_LOG || eax == IDC_EDITBOX_ABOUT_LOG
                Invoke SetBkMode,wParam,TRANSPARENT
                ;Invoke SetTextColor,wParam,COLOR_POLLO
                ;INVOKE SetBkColor,wParam,0 ;COLOR_POLLO
                mov eax,hPenPollo
                ret
            .else
                Invoke SetBkMode,wParam,TRANSPARENT  ;OPAQUE ;TRANSPARENT ;TRANSPARENT=Mancha(No actualiza)
                INVOKE SetBkColor,wParam,COLOR_NONE
                ;INVOKE SetBkColor,wParam,hBmpNoiseNormal
                ;Invoke CreatePatternBrush,hBmpNoiseNormal
                ret
          .endif
        ;.endif
      .elseif uMsg == WM_CTLCOLOREDIT ;WM_CTLCOLORDLG ;WM_CTLCOLORBTN
        ;Read-only or disabled edit controls do not send the WM_CTLCOLOREDIT message; instead, they send the WM_CTLCOLORSTATIC message.
        Invoke   GetDlgCtrlID, lParam
        .if eax == IDC_EDITBOX_HOTKEYS_SHOW_ME_LEFT || eax == IDC_EDITBOX_HOTKEYS_CAPTURE_TARGET || \
            eax == IDC_EDITBOX_HOTKEYS_FOR_EDIT_TARGET || eax == IDC_EDITBOX_HOTKEYS_CAPTURE_TOPWIN || \
            eax == IDC_EDITBOX_HOTKEYS_FOR_EDIT_TOPWIN || eax == IDC_EDITBOX_HOTKEYS_EMAIL_TARGET || \
            eax == IDC_EDITBOX_HOTKEYS_EMAIL_TOPWIN
            Invoke SetBkMode,wParam,TRANSPARENT
            ;Invoke SetTextColor,wParam,COLOR_POLLO
            ;INVOKE SetBkColor,wParam,0 ;COLOR_POLLO
            mov eax,hPenPollo
            ret
        .endif
    .elseif uMsg == WM_SIZE
        ;>>>>>>>>>>>>>> 5. TRAY ICON = Intercepting Minimize Event >>>>>>>>>>>>>
        ;
        .if wParam == SIZE_RESTORED
            .if dFirstTimeShowed==FALSE  ;First time showed is either by click on trayIcon or Hotkey
                Invoke ShowWindowAsync,hWinMain,SW_SHOWNORMAL ;SW_SHOW
            .else    
                mov dFirstTimeShowed,FALSE
                .if dIsWindowStyleTOOLWINDOW!=NULL  ;TODO: Remove TOOLWINDOW style when user changes settings checkbox for "hide window at..."
                    Invoke SetWindowLong,hWinMain,GWL_EXSTYLE,dIsWindowStyleTOOLWINDOW
                .endif
                Invoke ShowWindowAsync,hWinMain,SW_SHOWNORMAL ;SW_SHOW
            .endif
        .elseif wParam==SIZE_MINIMIZED
            Invoke IniManager,hWnd,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
            ;================================================
            ; Tray Icon
            ;================================================
            .if bShowANotificationIcon
                .if bIsShowANotificationIconActive==FALSE
                    Invoke Shell_NotifyIcon,NIM_ADD,addr TrayBalloonNote
                    ;Returns TRUE if successful, or FALSE otherwise.
                    mov bIsShowANotificationIconActive,al
                    .if dFirstTimeShowed==FALSE
                        Invoke ShowWindow,hWinMain,SW_HIDE
                    .endif
                .endif
            .else
                .if bIsShowANotificationIconActive==TRUE
                    Invoke Shell_NotifyIcon,NIM_DELETE,ADDR TrayBalloonNote
                    ;Returns TRUE if successful, or FALSE otherwise.
                    mov bIsShowANotificationIconActive,al
                .endif
            .endif
                ;================================================
                ; Hide Window
                ;================================================
            .if dFirstTimeShowed==FALSE
                .if bShowANotificationIcon && bIsShowANotificationIconActive
                    Invoke ShowWindow,hWinMain,SW_HIDE
                .elseif bHideApplicationWindowWhenOpening && bIsUnhideApplicationWindowHotKeyActive
                    Invoke ShowWindowAsync,hWinMain,SW_HIDE
                .else
                    Invoke ShowWindow,hWinMain,SW_MINIMIZE
                .endif
            .endif
            ;
        .endif
    ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ;>>>>>>>>>>>>>> 6. TRAY ICON = Intercepting WM_SHELLNOTIFY Event >>>>>>>>>>>>>
    .elseif uMsg==WM_SHELLNOTIFY ; Icon on system tray recieved a message
        .if wParam==IDM_TRAY_SHELLNOTIFY_MSG
            .if lParam==WM_LBUTTONDBLCLK
                mov dDoubleClick,2
                ;https://learn.microsoft.com/en-us/dotnet/desktop/winforms/input-mouse/how-to-distinguish-between-clicks-and-double-clicks?view=netdesktop-8.0
                ;SystemInformation.DoubleClickTime;
                ;TODO: Convertir código de IDM_TRAY_SHOW_APPLICATION en función y llamarla.
                Invoke IsIconic,hWinMain
                .if eax
                    Invoke ShowWindow,hWinMain,SW_RESTORE
                .endif
                Invoke IsWindowVisible,hWinMain
                .if eax
                    Invoke ShowWindow,hWinMain,SW_SHOW
                    Invoke SetForegroundWindow,hWinMain
                .endif           
            .elseif lParam==WM_LBUTTONDOWN
               ; Display popup menu to the correct alignment with the mouse position
;               mov dDoubleClick,1
;               mov eax,dDoubleClickSpeed
;               add eax,20
;               Invoke Sleep,eax               
;               Invoke IsWindowVisible,hWinMain
;               .if dDoubleClick==1 && eax
;                    ret
;               .endif 
                Invoke GetCursorPos, addr pnt
                Invoke SetForegroundWindow,hWnd
                Invoke TrackPopupMenu,hTrayPopupMenuLClick,TPM_BOTTOMALIGN+TPM_RIGHTALIGN+TPM_RETURNCMD, pnt.x,pnt.y,0,hWnd,0
                ;Invoke PostMessage,hWnd,WM_NULL,0,0
                .if ax==IDM_TRAY_SHOW_APPLICATION
                    Invoke IsIconic,hWinMain
                    .if eax
                        Invoke ShowWindow,hWinMain,SW_RESTORE
                    .endif
                    Invoke IsWindowVisible,hWinMain
                    .if eax
                        Invoke ShowWindow,hWinMain,SW_SHOW
                        Invoke SetForegroundWindow,hWinMain
                    .endif
                .elseif ax==IDM_TRAY_CAPTURE_SENT_TO_CLIPBOARD
                    Invoke StartAutoShotLoop,RUNTIME_CAPTURE_SENT_TO_CLIPBOARD,0,5,TRUE  ;bLocalTargetWindow,dSendArchiveTo,dMinutes,dSeconds,bOnlyOneTime
                .elseif ax==IDM_TRAY_CAPTURE_SENT_TO_LOCATION
                    Invoke StartAutoShotLoop,RUNTIME_CAPTURE_SENT_TO_LOCATION,0,5,TRUE
                .elseif ax==IDM_TRAY_CAPTURE_SENT_TO_MAIL
                    Invoke StartAutoShotLoop,RUNTIME_CAPTURE_SENT_TO_MAIL,0,5,TRUE
                .elseif ax==IDM_TRAY_CAPTURE_SENT_TO_OPEN
                    Invoke StartAutoShotLoop,RUNTIME_CAPTURE_SENT_TO_OPEN,0,5,TRUE
                .elseif ax == IDM_TRAY_QUIT_APPLICATION
                    jmp quit_dialog
                .endif
            .elseif lParam==WM_RBUTTONDOWN; Check if the user has right-clicked
                ; Display popup menu to the correct alignment with the mouse position
                Invoke GetCursorPos, addr pnt
                Invoke SetForegroundWindow,hWnd
                Invoke TrackPopupMenu,hTrayPopupMenuRClick,TPM_BOTTOMALIGN+TPM_RIGHTALIGN+TPM_RETURNCMD, pnt.x,pnt.y,0,hWnd,0
                .if ax==IDM_TRAY_SHOW_APPLICATION
                    Invoke IsIconic,hWinMain
                    .if eax
                        Invoke ShowWindow,hWinMain,SW_RESTORE
                    .endif
                    Invoke IsWindowVisible,hWinMain
                    .if eax
                        Invoke ShowWindow,hWinMain,SW_SHOW
                        Invoke SetForegroundWindow,hWinMain
                    .endif
                .elseif ax==IDM_TRAY_CAPTURE_SENT_TO_CLIPBOARD
                    Invoke StartAutoShotLoop,RUNTIME_CAPTURE_SENT_TO_CLIPBOARD,0,5,TRUE  ;bLocalTargetWindow,dSendArchiveTo,dMinutes,dSeconds,bOnlyOneTime
                .elseif ax==IDM_TRAY_CAPTURE_SENT_TO_LOCATION
                    Invoke StartAutoShotLoop,RUNTIME_CAPTURE_SENT_TO_LOCATION,0,5,TRUE
                .elseif ax==IDM_TRAY_CAPTURE_SENT_TO_MAIL
                    Invoke StartAutoShotLoop,RUNTIME_CAPTURE_SENT_TO_MAIL,0,5,TRUE
                .elseif ax==IDM_TRAY_CAPTURE_SENT_TO_OPEN
                    Invoke StartAutoShotLoop,RUNTIME_CAPTURE_SENT_TO_OPEN,0,5,TRUE
                .elseif ax == IDM_TRAY_QUIT_APPLICATION
                    jmp quit_dialog
                .endif   
            ;.elseif lParam == WM_MOUSEMOVE
            ;    invoke Shell_NotifyIcon,NIM_MODIFY,ADDR TrayBalloonNote
            ;    .if eax==FALSE
            ;        Invoke MessageBeep,NULL
            ;    .endif
            .endif
        .endif
;    .ELSEIF uMsg == 49316
;        invoke Shell_NotifyIcon,NIM_ADD,ADDR TrayBalloonNote
;                .if eax==FALSE
;                    Invoke MessageBeep,NULL
;                .endif
    ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    .elseif uMsg == WM_COMMAND
        xor ecx,ecx
        mov eax,wParam
        mov cx,ax
        .if ecx == IDC_DROP_CAPTURE_TARGET
            shr eax,16
            .if ((ax==CBN_CLOSEUP) || (ax==CBN_SELCHANGE))
                Invoke SendMessage, hCombo3, CB_GETCURSEL, 0, 0 ;The return value is the zero-based index of the currently selected item. If no item is selected, it is CB_ERR.
                .if eax==3 ;0=Monitor, 1=Active, 2=Desktop, 3=Region
                    Invoke GetDlgItem, hWnd,IDC_BTN_DIALOG_CAPTURE_REGION
                    Invoke EnableWindow,eax,TRUE
                .else
                    Invoke GetDlgItem, hWnd,IDC_BTN_DIALOG_CAPTURE_REGION
                    Invoke EnableWindow,eax,FALSE
                .endif
            .endif
        .elseif ecx == IDC_LIST_ITEM_BUTTON1
            Invoke IniManager,hWnd,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
            Invoke SelectPage,IDC_PAGE01
        .elseif ecx == IDC_LIST_ITEM_BUTTON2
            Invoke IniManager,hWnd,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
            Invoke SelectPage,IDC_PAGE02
        .elseif ecx == IDC_LIST_ITEM_BUTTON3
            Invoke IniManager,hWnd,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
            Invoke SelectPage,IDC_PAGE03
        .elseif ecx == IDC_LIST_ITEM_BUTTON4
            Invoke IniManager,hWnd,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)        
            Invoke SelectPage,IDC_PAGE04
        .elseif ecx == IDC_LIST_ITEM_BUTTON5
            Invoke IniManager,hWnd,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)        
            Invoke SelectPage,IDC_PAGE05
        .elseif ecx == IDC_LIST_ITEM_BUTTON6
            Invoke IniManager,hWnd,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)                      
            Invoke SelectPage,IDC_PAGE06
        .elseif ecx == IDC_LIST_ITEM_BUTTON7
            Invoke IniManager,hWnd,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
        .elseif ecx == IDC_LIST_ITEM_BUTTON8
            Invoke IniManager,hWnd,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)        
            jmp quit_dialog
        .elseif ecx == IDC_DROP_CAPTURE_MODE
            shr eax,16
            .if ((ax==CBN_CLOSEUP) || (ax==CBN_SELCHANGE))
                Invoke SendMessage, hCombo5, CB_GETCURSEL, 0, 0 ;The return value is the zero-based index of the currently selected item. If no item is selected, it is CB_ERR.
                push eax
                ;db 36 dup('Auto Shot',0,0,0,'Manual Shot',0,'Mause Click',0);12
                .if eax==0
                    .if bHookRunning==TRUE
                        Invoke UninstallHook
                        mov bHookRunning,FALSE
                    .endif
                    Invoke GetDlgItem, hWnd,IDC_PROMPT_TIMELAPSE
                    Invoke EnableWindow,eax,TRUE
                    Invoke GetDlgItem, hWnd,IDC_EDITBOX_TIMELAPSE_MIN
                    Invoke EnableWindow,eax,TRUE
                    Invoke GetDlgItem, hWnd,IDC_SPINBOX_TIMELAPSE_MIN
                    Invoke EnableWindow,eax,TRUE
                    Invoke GetDlgItem, hWnd,IDC_EDITBOX_TIMELAPSE_SEC
                    Invoke EnableWindow,eax,TRUE
                    Invoke GetDlgItem, hWnd,IDC_SPINBOX_TIMELAPSE_SEC
                    Invoke EnableWindow,eax,TRUE
                    Invoke GetDlgItem, hWnd,IDC_PROMPT_START_AUTOSHOT_AT_INIT
                    Invoke EnableWindow,eax,TRUE
                    Invoke GetDlgItem, hWnd,IDC_CHK_START_AUTOSHOT_AT_INIT
                    Invoke EnableWindow,eax,TRUE
                    Invoke GetDlgItem, hWnd,IDC_BTN_AUTOSHOT
                    Invoke EnableWindow,eax,TRUE
                    ;
                    Invoke GetDlgItem, hWnd,IDC_PROMPT_CLICK_DELAY
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_EDITBOX_CLICK_DELAY
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_SPINBOX_CLICK_DELAY
                    Invoke EnableWindow,eax,FALSE
                    ;
                .elseif eax==1
                    .if bHookRunning==TRUE
                        Invoke UninstallHook
                        mov bHookRunning,FALSE
                    .endif
                    Invoke GetDlgItem, hWnd,IDC_PROMPT_TIMELAPSE
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_EDITBOX_TIMELAPSE_MIN
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_SPINBOX_TIMELAPSE_MIN
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_EDITBOX_TIMELAPSE_SEC
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_SPINBOX_TIMELAPSE_SEC
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_PROMPT_START_AUTOSHOT_AT_INIT
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_CHK_START_AUTOSHOT_AT_INIT
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_BTN_AUTOSHOT
                    Invoke EnableWindow,eax,FALSE
                    ;
                    Invoke GetDlgItem, hWnd,IDC_PROMPT_CLICK_DELAY
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_EDITBOX_CLICK_DELAY
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_SPINBOX_CLICK_DELAY
                    Invoke EnableWindow,eax,FALSE
                .elseif eax==2
                    ;********* CaptureMode
                    .if bHookRunning==FALSE
                        Invoke InstallHook,hWinMain,WM_MOUSEHOOK,NULL
                        .if eax==NULL
                            mov bHookRunning,FALSE
                        .else
                            mov bHookRunning,TRUE
                        .endif
                    .endif
                    ;
                    Invoke GetDlgItem, hWnd,IDC_PROMPT_TIMELAPSE
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_EDITBOX_TIMELAPSE_MIN
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_SPINBOX_TIMELAPSE_MIN
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_EDITBOX_TIMELAPSE_SEC
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_SPINBOX_TIMELAPSE_SEC
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_PROMPT_START_AUTOSHOT_AT_INIT
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_CHK_START_AUTOSHOT_AT_INIT
                    Invoke EnableWindow,eax,FALSE
                    Invoke GetDlgItem, hWnd,IDC_BTN_AUTOSHOT
                    Invoke EnableWindow,eax,FALSE
                    ;
                    Invoke GetDlgItem, hWnd,IDC_PROMPT_CLICK_DELAY
                    Invoke EnableWindow,eax,TRUE
                    Invoke GetDlgItem, hWnd,IDC_EDITBOX_CLICK_DELAY
                    Invoke EnableWindow,eax,TRUE
                    Invoke GetDlgItem, hWnd,IDC_SPINBOX_CLICK_DELAY
                    Invoke EnableWindow,eax,TRUE
                .else
                    .if bHookRunning==TRUE
                        Invoke UninstallHook
                        mov bHookRunning,FALSE
                    .endif
                .endif
            .endif
        .elseif wParam == IDC_BTN_DIALOG_ARCHIVE_PATH                      ;Destination Path
            shr eax,16
            .if ax==BN_CLICKED
                mov szBuffer,0
                Invoke GetDlgItemText,hWnd,IDC_EDITBOX_ARCHIVE_PATH,addr szBuffer,MAX_PATH
                Invoke FillBuffer,ADDR szBuffer,length szBuffer,0
                Invoke GetFileName,hWnd,ADDR szBuffer,addr szNoFilter
                cmp szBuffer[0],0  ;<< zero if cancel pressed in dlgbox
                je @F
                    Invoke SetDlgItemText,hWnd,IDC_EDITBOX_ARCHIVE_PATH,addr szBuffer
                @@:
            .endif
;        .elseif wParam == IDC_BTN_HOTKEY                  ;HotKeys Settings
;            shr eax,16
;            .if ax==BN_CLICKED
;                Invoke HotKeysDlg,hWnd,hIcon
;            .endif
        .elseif wParam == IDC_BTN_DIALOG_CAPTURE_REGION               ;Region Setting
            shr eax,16
            .if ax==BN_CLICKED
                Invoke RegionDlg,hWnd,hIcon
            .endif
        .elseif wParam == IDC_BTN_SEND_TO_LOCATION ;RUNTIME_CAPTURE_SENT_TO_LOCATION
            shr eax,16
            .if ax==BN_CLICKED
                Invoke StartAutoShotLoop,RUNTIME_CAPTURE_SENT_TO_LOCATION,0,5,TRUE ;bLocalTargetWindow,dSendArchiveTo,dMinutes,dSeconds,bOnlyOneTime
            .endif
        .elseif wParam == IDC_BTN_SEND_TO_EMAIL
            shr eax,16
            .if ax==BN_CLICKED
                Invoke StartAutoShotLoop,RUNTIME_CAPTURE_SENT_TO_MAIL,0,5,TRUE
            .endif
        .elseif wParam == IDC_BTN_SEND_TO_CLIPBOARD
            shr eax,16
            .if ax==BN_CLICKED
                Invoke StartAutoShotLoop,RUNTIME_CAPTURE_SENT_TO_CLIPBOARD,0,5,TRUE
            .endif
        .elseif wParam == IDC_BTN_AUTOSHOT
            shr eax,16
            .if ax==BN_CLICKED
                ;
                .if dAutoShotLoop_IsRunning==TRUE
                    Invoke StopAutoShotLoop
                .else
                    Invoke StartAutoShotLoop,NULL,bAutoShot_TimerMinutes,bAutoShot_TimerSeconds,FALSE
                .endif
            .endif
        .elseif wParam == IDC_BTN_HIDE                  ;Cerrar
            shr eax,16
            .if ax==BN_CLICKED
                Invoke IniManager,hWnd,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER) ;Button "_" & "X"
                Invoke ShowWindow,hWnd,SW_MINIMIZE
            .endif
        .endif
    .elseif uMsg == WM_SYSCOMMAND
        and wParam,0FFFFFFF0H
        .if (wParam==SC_MINIMIZE) ;SC_ICON = SC_MINIMIZE Cerro desde la X del SystemMenu
            Invoke IniManager,hWnd,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER) ;Button "_" & "X"
;            .if bCaptureMode==RUNTIME_CAPTURE_MODE_AUTOSHOT
;                Invoke SendMessage,hWinMain, WM_COMMAND, IDC_BTN_AUTOSHOT, 0
;            .endif
        .elseif (wParam == SC_CLOSE)
            Invoke IniManager,hWnd,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
                mov bExeCloseIsAllowed,FALSE
                ;if dRunningCapture == FALSE
                ;   Invoke SendMessage,hWinMain,WM_COMMAND,IDC_BTN_AUTOSHOT,0
                ;endif
                Invoke ShowWindow,hWnd,SW_MINIMIZE
        ;.elseif wParam == SC_SCREENSAVE   ;SC_MOUSEMENU SC_MONITORPOWER
        .endif
    .elseif uMsg == WM_CLOSE
        .if bExeCloseIsAllowed==FALSE
            mov bExeCloseIsAllowed,TRUE
            jmp NextLoop
        .endif
        quit_dialog:
        mov bExeIsClosing,TRUE
        .if bHookRunning==TRUE
            Invoke Sleep,125
            Invoke UninstallHook
            mov bHookRunning,FALSE
            Invoke Sleep,125
        .endif
        Invoke DestroyMenu,hMacroPopupMenu
        ;********* Remove tray icon, before close
    ;>>>>>>>>>>>>>> 7. TRAY ICON = Remove tray icon, before close >>>>>>>>>>>>>
        Invoke Shell_NotifyIcon, NIM_DELETE,addr TrayBalloonNote
        Invoke DestroyMenu,hTrayPopupMenuRClick
        Invoke DestroyMenu,hTrayPopupMenuLClick
    ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ;
         .if dAutoShotLoop_IsRunning == TRUE || hAutoShotLoop_ThreadHandle!=NULL
            Invoke CloseHandle,hAutoShotLoop_ThreadHandle
        .endif
        ;
        Invoke GlobalFindAtom,addr szAtom
        .if eax
            Invoke GlobalDeleteAtom,eax
        .endif
        Invoke GlobalDeleteAtom,hSingleInstanceAtom
        ;
        Invoke CapturA_FreeLibrary
        ;
        Invoke DeleteObject,hPenPollo
        Invoke DeleteObject,LightMode.hTitleFont
        Invoke DeleteObject,LightMode.hSubTitleFont
        Invoke DeleteObject,LightMode.hRegularFont
        Invoke EndDialog,hWnd,0
        ;
    .elseif uMsg==WM_HOTKEY
        .if wParam==HOTKEY_PRINTSCREEN
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            ;szText szActivity_HotKey_PrintScreen,"New hotkey message: PRINTSCREEN"
            ;Invoke UpdateActivityLog,addr szActivity_HotKey_PrintScreen,NULL,1,1
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ 
            Invoke TakeScreenShot,bCaptureWindow,RUNTIME_CAPTURE_SENT_TO_NONE
        .elseif wParam==HOTKEY_CTRL_PRINTSCREEN
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            ;szText szActivity_HotKey_CtrlPrintScreen,"New hotkey message: CTRL+PRINTSCREEN"
            ;Invoke UpdateActivityLog,addr szActivity_HotKey_CtrlPrintScreen,NULL,1,1
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke TakeScreenShot,bCaptureWindow,RUNTIME_CAPTURE_SENT_TO_OPEN
        .elseif wParam==HOTKEY_ALT_PRINTSCREEN
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            ;szText szActivity_HotKey_AltPrintScreen,"New hotkey message: ALT+PRINTSCREEN"
            ;Invoke UpdateActivityLog,addr szActivity_HotKey_AltPrintScreen,NULL,1,1
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke TakeScreenShot,RUNTIME_TARGET_TYPE_ACTIVE_WINDOWS,RUNTIME_CAPTURE_SENT_TO_NONE
        .elseif wParam==HOTKEY_CTRL_ALT_PRINTSCREEN
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            ;szText szActivity_HotKey_CtrlAltPrintScreen,"New hotkey message: CTRL+ALT+PRINTSCREEN"
            ;Invoke UpdateActivityLog,addr szActivity_HotKey_CtrlAltPrintScreen,NULL,1,1
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke TakeScreenShot,RUNTIME_TARGET_TYPE_ACTIVE_WINDOWS,RUNTIME_CAPTURE_SENT_TO_OPEN
        .elseif wParam==HOTKEY_SHIFT_CTRL_E
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            ;szText szActivity_HotKey_ShiftCtrlE,"New hotkey message: SHIFT+CTRL+E"
            ;Invoke UpdateActivityLog,addr szActivity_HotKey_ShiftCtrlE,NULL,1,1
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke TakeScreenShot,bCaptureWindow,RUNTIME_CAPTURE_SENT_TO_MAIL
        .elseif wParam==HOTKEY_SHIFT_CTRL_ALT_E
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            ;szText szActivity_HotKey_ShiftCtrlAltE,"New hotkey message: SHIFT+CTRL+ALT+E"
            ;Invoke UpdateActivityLog,addr szActivity_HotKey_ShiftCtrlAltE,NULL,1,1
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke TakeScreenShot,RUNTIME_TARGET_TYPE_ACTIVE_WINDOWS,RUNTIME_CAPTURE_SENT_TO_MAIL
        .elseif wParam==HOTKEY_SHIFT_CTRL_ALT_PRINTSCREEN || HOTKEY_SHIFT_CTRL_ALT_S
            .if wParam==HOTKEY_SHIFT_CTRL_ALT_PRINTSCREEN
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                szText szActivity_HotKey_ShiftCtrlAltPrintScreen,"-New hotkey message: SHIFT+CTRL+ALT+PRINTSCREEN"
                Invoke UpdateActivityLog,addr szActivity_HotKey_ShiftCtrlAltPrintScreen,NULL,1,0
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ 
            .else
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                szText szActivity_HotKey_ShiftCtrlAltS,"-New hotkey message: SHIFT+CTRL+ALT+S"
                Invoke UpdateActivityLog,addr szActivity_HotKey_ShiftCtrlAltS,NULL,1,0
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ 
            .endif
            Invoke IsIconic,hWinMain
            .if eax
                Invoke ShowWindow,hWinMain,SW_RESTORE
            .endif
            Invoke IsWindowVisible,hWinMain
            .if eax
                Invoke ShowWindow,hWinMain,SW_SHOW
                Invoke SetForegroundWindow,hWinMain
            .endif
        .endif
    .elseif uMsg==WM_MOUSEHOOK
        ;.if wParam == WM_LBUTTONDOWN || wParam == WM_NCLBUTTONDOWN || WM_RBUTTONDOWN || wParam == WM_NCRBUTTONDOWN
            Invoke TakeScreenShot,bCaptureWindow,RUNTIME_CAPTURE_SENT_TO_NONE
        ;.endif
;    .else
;        invoke DefWindowProc, hWnd, uMsg, wParam, lParam   ; default processing
;        ret
    .endif
    ;Invoke GetAsyncKeyState,VK_SNAPSHOT ;VK_SHIFT+, VK_CONTROL, and VK_MENU
    NextLoop:
    xor eax, eax
    ret
WinMainProc                     endp
align 4
InitDialogWinMain               proc hWnd:DWORD
                                LOCAL Rct:RECT
        ;
        xor ecx,ecx
        mov dRunningCapture,ecx
        mov hAutoShotLoop_ThreadHandle,eax
        ;
        Invoke CreateSolidBrush,COLOR_POLLO
        mov hPenPollo,eax

        mov eax,hWnd
        mov hWinMain,eax
        mov hOwnerPopMenu,eax
        ;
        mov bExeCloseIsAllowed,TRUE
        mov bExeIsClosing,FALSE
        ;
        Invoke UpdateActivityLog,-1,-1,0,0 ;Update WinMain EntryPoint
        szText szActivity_InitDialog,9,9,9,"***Init Window Dialog***"
        Invoke UpdateActivityLog,addr szActivity_InitDialog,NULL,2,0 ;First time called with hWinMain
        ;
        Invoke GetDlgItem,hWnd,IDC_LIST_GROUP_OPTIONS
        Invoke ShowWindow,eax,SW_HIDE
        ;
        Invoke GetDlgItem,hWnd,IDC_LIST_ITEM_BUTTON1
        Invoke SendMessage,eax,BM_SETCHECK,BST_CHECKED,NULL
        ;
        Invoke GetDlgItem,hWnd,IDC_LIST_ITEM_BUTTON7
        Invoke ShowWindow,eax,SW_HIDE ;Disabled in SelectPage

;        Invoke GetDlgItem,hWnd,IDC_SHEET_TAB_CONTROL_COVER
;        push eax
;        Invoke SetWindowPos,eax, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE OR SWP_NOSIZE ;To mov Cover to top, all other controsl must have "WS_CLIPSIBLINGS"
;        pop eax
;        Invoke ShowWindow,eax,SW_HIDE
        ;
        ;Put File Version in Windows Title
        Invoke GetFileVersionInfoSize,OFFSET szProgramPath,OFFSET dDummy
        .if (eax != 0)
              push eax
              Invoke GlobalAlloc,GPTR,eax
              mov pMemory,eax
              pop eax
              Invoke GetFileVersionInfo,OFFSET szProgramPath,NULL,eax,pMemory
              Invoke VerQueryValue,pMemory,OFFSET szFileVersion,OFFSET dDummy,OFFSET dwVersionLen
              Invoke GetWindowText,hWnd,addr szBuffer,MAX_PATH
                  szText szVerTitle,", version "
                  Invoke lstrcat,addr szBuffer,addr szVerTitle
                  Invoke lstrcat,addr szBuffer,addr szVer
                  Invoke lstrcat,addr szBuffer,dDummy
                  ;Invoke lstrcat,addr szBuffer,addr szBracketOpen
                  ;Invoke lstrcat,addr szBuffer,addr szProgramPath
                  ;Invoke lstrcat,addr szBuffer,addr szBracketClose
              Invoke SetWindowText,hWnd,addr szBuffer
              Invoke GlobalFree,pMemory
        .endif
        ;
       ;
        ;Invoke CreateSolidBrush,COLOR_WINDOW+1
        ;mov hBkbrush,eax
        ;;Invoke SetClassLong,hWnd,GCL_HBRBACKGROUND,hBkbrush
        ;Invoke SetClassLong,hWnd,GCL_HBRBACKGROUND,hBkbrush
        ;;Invoke SetWindowLongPtr,hWnd,GCL_HBRBACKGROUND,hBkbrush
        ;;InvalidateRect(hwnd, NULL, TRUE);
        ;Invoke SetWindowPos,hWnd,NULL,0,0,0,0,SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOZORDER OR SWP_FRAMECHANGED
        ;
        ;
        ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        ;>>>>>>>>>>>>>> 4. TRAY ICON = Create PopUpMenu >>>>>>>>>>>>>
        ;
        szText szTrayShowApplication,"CapturA+ settings"
        szText szTrayWaitAndSendToClipBoard,"Wait 5 secs to snap && copy to clipboad"
        szText szTrayWaitAndSendToLocation,"Wait 5 secs to snap && open location"
        szText szTrayWaitAndSendToMail,"Wait 5 secs to snap && open email client"
        szText szTrayWaitAndSendToOpen,"Wait 5 secs to snap && open default editor"
        szText szTrayQuitApplication,"Quit application"
        
        Invoke CreatePopupMenu
        mov hTrayPopupMenuLClick,eax
        ;Invoke AppendMenu,hTrayPopupMenuLClick,MF_STRING,IDM_TRAY_SHOW_APPLICATION,addr szTrayShowApplication
        ;Invoke AppendMenu,hTrayPopupMenuLClick,MF_SEPARATOR, 0, 0
        Invoke AppendMenu,hTrayPopupMenuLClick,MF_STRING,IDM_TRAY_CAPTURE_SENT_TO_CLIPBOARD,addr szTrayWaitAndSendToClipBoard
        Invoke AppendMenu,hTrayPopupMenuLClick,MF_STRING,IDM_TRAY_CAPTURE_SENT_TO_LOCATION,addr szTrayWaitAndSendToLocation
        Invoke AppendMenu,hTrayPopupMenuLClick,MF_STRING,IDM_TRAY_CAPTURE_SENT_TO_MAIL,addr szTrayWaitAndSendToMail
        Invoke AppendMenu,hTrayPopupMenuLClick,MF_STRING,IDM_TRAY_CAPTURE_SENT_TO_OPEN,addr szTrayWaitAndSendToOpen
        ;Invoke AppendMenu,hTrayPopupMenuLClick,MF_SEPARATOR, 0, 0
        ;Invoke AppendMenu,hTrayPopupMenuLClick,MF_STRING,IDM_TRAY_QUIT_APPLICATION,addr szTrayQuitApplication
        ;
        Invoke CreatePopupMenu
        mov hTrayPopupMenuRClick,eax
        Invoke AppendMenu,hTrayPopupMenuRClick,MF_STRING,IDM_TRAY_SHOW_APPLICATION,addr szTrayShowApplication
        Invoke AppendMenu,hTrayPopupMenuRClick,MF_SEPARATOR, 0, 0
        Invoke AppendMenu,hTrayPopupMenuRClick,MF_STRING,IDM_TRAY_CAPTURE_SENT_TO_CLIPBOARD,addr szTrayWaitAndSendToClipBoard
        Invoke AppendMenu,hTrayPopupMenuRClick,MF_STRING,IDM_TRAY_CAPTURE_SENT_TO_LOCATION,addr szTrayWaitAndSendToLocation
        Invoke AppendMenu,hTrayPopupMenuRClick,MF_STRING,IDM_TRAY_CAPTURE_SENT_TO_MAIL,addr szTrayWaitAndSendToMail
        Invoke AppendMenu,hTrayPopupMenuRClick,MF_STRING,IDM_TRAY_CAPTURE_SENT_TO_OPEN,addr szTrayWaitAndSendToOpen
        Invoke AppendMenu,hTrayPopupMenuRClick,MF_SEPARATOR, 0, 0
        Invoke AppendMenu,hTrayPopupMenuRClick,MF_STRING,IDM_TRAY_QUIT_APPLICATION,addr szTrayQuitApplication
        ;
        Invoke GetWindowText,hWinMain,addr szBuffer,MAX_PATH
        ;
        ;For _WIN32_WINNT  > 0x0501, you must set 'cbSize' to TTTOOLINFOA_V2_SIZE (instead of sizeof(TOOLTIPINFO))
        ;or include the appropriate version of Common Controls in the manifest. Otherwise the tooltip won't be displayed.
        mov TrayBalloonNote.cbSize,sizeof NOTIFYICONDATA ;sizeof TrayBalloonNote.szTip ;sizeof NOTIFYICONDATA ;sizeof TrayBalloonNote.szTip ;TTTOOLINFOA_V1_SIZE,NOTIFYICONDATA
        m2m TrayBalloonNote.hwnd, hWnd
        mov TrayBalloonNote.uID,IDM_TRAY_SHELLNOTIFY_MSG
        mov TrayBalloonNote.uCallbackMessage,WM_SHELLNOTIFY
        mov TrayBalloonNote.uFlags,NIF_ICON+NIF_MESSAGE+NIF_TIP+NIF_INFO
        mov TrayBalloonNote.uCallbackMessage,WM_SHELLNOTIFY
        m2m TrayBalloonNote.hIcon, hIconTray; hIconAboutUnpush ; hIconTray
        Invoke lstrcpy,addr TrayBalloonNote.szTip,addr szAppTitle
        mov TrayBalloonNote.uVersion,NIM_SETVERSION ;Menor a Win2000 = NIM_SETVERSION, Mayor o igual win2000 = NOTIFYICON_VERSION
        mov TrayBalloonNote.dwInfoFlags,NIIF_WARNING
        ;Invoke SetWindowText,hWinMain,ADDR szAppName
        ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        ;
        ;CREATE TOOLTIP
        Invoke CreateWindowEx,NULL,addr szTOOLTIPS_CLASS,NULL,\
                                WS_POPUP OR TTS_NOPREFIX OR TTS_ALWAYSTIP OR TTS_BALLOON\
                                ,0,0,0,0,hWnd,NULL,hInstance,NULL
        mov    hToolTip,eax
        Invoke SetWindowPos,hToolTip,HWND_TOPMOST,0,0,0,0,SWP_NOMOVE or SWP_NOSIZE or SWP_NOACTIVATE
        Invoke SendMessage,hToolTip,TTM_ACTIVATE,TRUE,0
        ;INITIALIZE MEMBERS OF THE TOOLINFO STRUCTURE
        mov    ti.cbSize,sizeof TOOLINFO
        mov    ti.uFlags,TTF_SUBCLASS or TTF_IDISHWND
        push   hWnd
        pop    ti.hWnd
        push   hInstance
        pop    ti.hInst
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDITBOX_ARCHIVE_PATH
        Invoke GetDlgItem, hWnd,IDC_EDITBOX_ARCHIVE_PATH
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTxt0
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDITBOX_JPG_COMMENTS
        Invoke GetDlgItem, hWnd,IDC_EDITBOX_JPG_COMMENTS
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTxt0
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDITBOX_TIMELAPSE_MIN
        Invoke GetDlgItem, hWnd,IDC_EDITBOX_TIMELAPSE_MIN
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTxt1
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDITBOX_TIMELAPSE_SEC
        Invoke GetDlgItem, hWnd,IDC_EDITBOX_TIMELAPSE_SEC
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTxt2
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDITBOX_IMAGE_QUALITY
        Invoke GetDlgItem, hWnd,IDC_EDITBOX_IMAGE_QUALITY
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTxt3
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDITBOX_COUNTER
        Invoke GetDlgItem, hWnd,IDC_EDITBOX_COUNTER
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTxt4
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_DROP_CAPTURE_TARGET
        Invoke GetDlgItem, hWnd,IDC_DROP_CAPTURE_TARGET
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTxt5
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        Invoke SendMessage,hToolTip,TTM_SETMAXTIPWIDTH,0,400
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_DROP_CAPTURE_MODE
        Invoke GetDlgItem, hWnd,IDC_DROP_CAPTURE_MODE
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTipTxt_Mode
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        Invoke SendMessage,hToolTip,TTM_SETMAXTIPWIDTH,0,400
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_DROP_CAPTURE_CURSOR
        Invoke GetDlgItem, hWnd,IDC_DROP_CAPTURE_CURSOR
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTipTxt_Cursor
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        Invoke SendMessage,hToolTip,TTM_SETMAXTIPWIDTH,0,400
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDITBOX_CLICK_DELAY
        ;Invoke GetRegistryMouseMenu,addr szMouseHoverTime,addr szMenuShowDelay ; Called within at app start
        Invoke lstrcpy,addr szToolTipTxt_Delay_Solved,addr szToolTipTxt_Delay1
        Invoke lstrcat,addr szToolTipTxt_Delay_Solved,addr szMouseHoverTime
        Invoke lstrcat,addr szToolTipTxt_Delay_Solved,addr szToolTipTxt_Delay2
        Invoke lstrcat,addr szToolTipTxt_Delay_Solved,addr szMenuShowDelay
        Invoke lstrcat,addr szToolTipTxt_Delay_Solved,addr szToolTipTxt_Delay3
        Invoke GetDlgItem, hWnd,IDC_EDITBOX_CLICK_DELAY
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTipTxt_Delay_Solved
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        Invoke SendMessage,hToolTip,TTM_SETMAXTIPWIDTH,0,400
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDITBOX_IMAGE_ZOOM
        Invoke GetDlgItem, hWnd,IDC_EDITBOX_IMAGE_ZOOM
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTipTxt_Zoom
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        Invoke SendMessage,hToolTip,TTM_SETMAXTIPWIDTH,0,400
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHK_HIDE_WIN_WHEN_OPENING
        Invoke GetDlgItem, hWnd,IDC_CHK_HIDE_WIN_WHEN_OPENING
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTxt8
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        Invoke SendMessage,hToolTip,TTM_SETMAXTIPWIDTH,0,400
        ;
        ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHK_HIDE_WIN_WHEN_OPENING
        Invoke GetDlgItem, hWnd,IDC_BTN_HIDE
        mov    ti.uId,eax
        mov    ti.lpszText,offset szToolTxt8
        Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
        Invoke SendMessage,hToolTip,TTM_SETMAXTIPWIDTH,0,400
        ;
        ;
        ;*************** Fills up Target Type  List ******************
        ;
        Invoke GetDlgItem, hWnd, IDC_DROP_TARGET_FILE_FORMAT
        mov hCombo1, eax
        mov eax,0
        @@:
        mov dCounter,eax
        lea ebx,szFileType
        add ebx,eax
        Invoke SendMessage, hCombo1, CB_ADDSTRING, 0, ebx
        mov eax,dCounter
        add eax,40       ;szFileType db 40 dup('JGP File (*.jpg)',0,0,0,0,'PNG File (*.png)',0,0,0,0);20
        cmp eax,20       ;2 arrays items. Only show the first one.
        jle @B
        ;
        movzx eax,bTargetType
        dec eax
        Invoke SendMessage, hCombo1, CB_SETCURSEL, eax, 0
        ;
;        ;*************** Fills up Store Mode  List ******************
;        ;
;        Invoke GetDlgItem, hWnd, IDC_DROP_ARCHIVE_MODE
;        mov hCombo2, eax
;        mov eax,0
;        @@:
;        mov dCounter,eax
;        lea ebx,szStoreMode
;        add ebx,eax
;        Invoke SendMessage, hCombo2, CB_ADDSTRING, 0, ebx
;        mov eax,dCounter
;        add eax,15       ;szStoreMode      db 2 dup('Single File',0,0,0,0,'Append to File',0);15
;        cmp eax,15       ;Dos elementos del array
;        jle @B
;        ;
;        movzx eax,bArchiveMode
;        dec eax
;        Invoke SendMessage, hCombo2, CB_SETCURSEL, eax, 0
        ;
        ;*************** Fills up ArchivePath  Box ******************
        ;
        Invoke SetDlgItemText,hWnd,IDC_EDITBOX_ARCHIVE_PATH,addr szArchivePath
        ;
        ;*************** Fills up JpgComments  Box ******************
        ;
        Invoke SetDlgItemText,hWnd,IDC_EDITBOX_JPG_COMMENTS,addr szJpgComments
        ;
        ;*************** Fills up Capture Window List ******************
        ;
        Invoke GetDlgItem, hWnd, IDC_DROP_CAPTURE_TARGET
        mov hCombo3, eax
        mov eax,0
        @@:
        mov dCounter,eax
        lea ebx,szCaptureWindow
        add ebx,eax
        Invoke SendMessage, hCombo3, CB_ADDSTRING, 0, ebx
        mov eax,dCounter
        add eax,15           ;szCaptureWindow             db 60 dup('Active Monitor',0,'Active Window',0,0,'DeskTop',0,0,0,0,0,0,0,0,'Region',0,0,0,0,0,0,0,0,0);15
        cmp eax,45          ;Cuatro elementos del array
        jle @B
        ;
        movzx eax,bCaptureWindow
        dec eax
        Invoke SendMessage,hCombo3,CB_SETCURSEL, eax, 0
        .if bCaptureWindow==3 ;0=Monitor, 1=Active, 2=Desktop, 3=Region
            Invoke GetDlgItem,hWnd,IDC_BTN_DIALOG_CAPTURE_REGION
            Invoke EnableWindow,eax,TRUE
        .else
            Invoke GetDlgItem,hWnd,IDC_BTN_DIALOG_CAPTURE_REGION
            Invoke EnableWindow,eax,FALSE
        .endif
        ;
        ;*************** Fills up Capture Cursor  List ******************
        ;
        Invoke GetDlgItem, hWnd, IDC_DROP_CAPTURE_CURSOR
        mov hCombo4, eax
        Invoke SendMessage, hCombo4, CB_ADDSTRING, 0, addr szCaptureCursor
        lea ecx, szCaptureCursor
        add ecx,15                   ;db 45 dup('Disable',0,0,0,0,0,0,0,0,'Default Cursor',0,'Red Arrow',0,0,0,0,0,0);15
        Invoke SendMessage, hCombo4, CB_ADDSTRING, 0, ecx
         .if hCursor
            mov eax,hCursor
            mov hCursorRed,eax
            mov hCursor,NULL ;Default System Arrow cursor
            lea ecx, szCursorFileName
        .else
            lea ecx, szCaptureCursor
            add ecx,30   ;szCaptureCursor             db 45 dup('Disable',0,0,0,0,0,0,0,0,'Default Cursor',0,'Red Arrow',0,0,0,0,0,0);15
        .endif
        Invoke SendMessage, hCombo4, CB_ADDSTRING, 0, ecx
        ;

        ;
        movzx eax,bCaptureCursor
        dec eax
        Invoke SendMessage, hCombo4, CB_SETCURSEL, eax, 0
        ;
        ;*************** Fills up Image Mode  List ******************
        ;
        Invoke GetDlgItem, hWnd, IDC_DROP_COLOR_IMAGE_MODE
        mov hCombo6, eax
        mov eax,0
        @@:
        mov dCounter,eax
        lea ebx,szImageMode
        add ebx,eax
        Invoke SendMessage, hCombo6, CB_ADDSTRING, 0, ebx
        mov eax,dCounter
        add eax,11       ;db 2 dup('Color',0,0,0,0,0,0'Gray Scale',0,);11
        cmp eax,11       ;Dos elementos del array
        jle @B
        ;
        movzx eax,bImageMode
        dec eax
        Invoke SendMessage, hCombo6, CB_SETCURSEL, eax, 0
        ;
        ;
        ;*************** Fills up Capture Mode  List ******************
        ;
        Invoke GetDlgItem,hWnd,IDC_DROP_CAPTURE_MODE
        mov hCombo5, eax
        mov eax,0
        @@:
        mov dCounter,eax
        lea ebx,szCaptureMode
        add ebx,eax
        Invoke SendMessage,hCombo5,CB_ADDSTRING,0,ebx ;wParam=This parameter is not used,lParam=An LPCTSTR pointer to the null-terminated string to be added.
        mov eax,dCounter
        add eax,12       ;db 36 dup('Auto Shot',0,0,0,'Manual Shot',0,'Mause Click',0);12
        cmp eax,24       ;Tres elementos del array
        jle @B
        ;
        Invoke CapturA_LoadLibrary
        .if eax==NULL
            Invoke SendMessage,hCombo5,CB_DELETESTRING,2,0 ;wParam=The zero-based index of the string to delete,lParam=This parameter is not used.
        .endif
        ;
        movzx eax,bCaptureMode
        dec eax
        Invoke SendMessage, hCombo5, CB_SETCURSEL, eax, 0
        .if bCaptureMode!=RUNTIME_CAPTURE_MODE_AUTOSHOT ;CAPTURE MODE: RUNTIME_CAPTURE_MODE_AUTOSHOT,RUNTIME_CAPTURE_MODE_MANUAL,RUNTIME_CAPTURE_MODE_CLICKS
            Invoke GetDlgItem, hWnd,IDC_PROMPT_TIMELAPSE
            Invoke EnableWindow,eax,FALSE
            Invoke GetDlgItem, hWnd,IDC_EDITBOX_TIMELAPSE_MIN
            Invoke EnableWindow,eax,FALSE
            Invoke GetDlgItem, hWnd,IDC_SPINBOX_TIMELAPSE_MIN
            Invoke EnableWindow,eax,FALSE
            Invoke GetDlgItem, hWnd,IDC_EDITBOX_TIMELAPSE_SEC
            Invoke EnableWindow,eax,FALSE
            Invoke GetDlgItem, hWnd,IDC_SPINBOX_TIMELAPSE_SEC
            Invoke EnableWindow,eax,FALSE
            Invoke GetDlgItem, hWnd,IDC_PROMPT_START_AUTOSHOT_AT_INIT
            Invoke EnableWindow,eax,FALSE
            Invoke GetDlgItem, hWnd,IDC_CHK_START_AUTOSHOT_AT_INIT
            Invoke EnableWindow,eax,FALSE
            Invoke GetDlgItem, hWnd,IDC_BTN_AUTOSHOT
            Invoke EnableWindow,eax,FALSE
        .endif
        ;
        .if bCaptureMode==RUNTIME_CAPTURE_MODE_MANUAL
            Invoke GetDlgItem, hWnd,IDC_PROMPT_CLICK_DELAY
            Invoke EnableWindow,eax,FALSE
            Invoke GetDlgItem, hWnd,IDC_EDITBOX_CLICK_DELAY
            Invoke EnableWindow,eax,FALSE
            Invoke GetDlgItem, hWnd,IDC_SPINBOX_CLICK_DELAY
            Invoke EnableWindow,eax,FALSE
        .endif
        ;
        ;
        ;*************** Fills up Mouse Delay (ms) ******************
        ;
        mov eax,dMouseClickSnapDelay
        Invoke SetDlgItemInt,hWnd,IDC_EDITBOX_CLICK_DELAY,eax,FALSE
        ;
        Invoke  GetDlgItem, hWnd, IDC_SPINBOX_CLICK_DELAY
        Invoke  SendMessage, eax, UDM_SETRANGE, 0, MAX_CLICK_DELAY ;000007D0H   ;wordLower,wordUpper 0,2000
        ;
        ;
        ;*************** Fills up TimerMinutes ******************
        ;
        movzx eax,bAutoShot_TimerMinutes
        Invoke SetDlgItemInt,hWnd,IDC_EDITBOX_TIMELAPSE_MIN,eax,FALSE
        ;
        Invoke  GetDlgItem, hWnd, IDC_SPINBOX_TIMELAPSE_MIN
        Invoke  SendMessage, eax, UDM_SETRANGE, 0, 00000063H   ;wordLower,wordUpper 0,99
        ;
        ;*************** Fills up TimerSeconds ******************
        ;
        movzx eax,bAutoShot_TimerSeconds
        Invoke SetDlgItemInt,hWnd,IDC_EDITBOX_TIMELAPSE_SEC,eax,FALSE
        ;
        Invoke  GetDlgItem, hWnd, IDC_SPINBOX_TIMELAPSE_SEC
        Invoke  SendMessage, eax, UDM_SETRANGE, 0, 0000003BH   ;wordLower,wordUpper 0,59
        ;
      ;
        ;*************** Initialize Check Control for bAutoShotStartAtInit  ******************
        ;
        Invoke GetDlgItem,hWnd,IDC_CHK_START_AUTOSHOT_AT_INIT
        .if bAutoShotStartAtInit
            mov ecx,BST_CHECKED
        .else
            mov ecx,BST_UNCHECKED
        .endif
        Invoke  SendMessage, eax, BM_SETCHECK,ecx,NULL
        ;
        ;
        ;*************** Initialize Check Control for bCopyFilePathToClipboard ******************
        ;
        Invoke GetDlgItem,hWnd,IDC_CHK_COPY_PATH_TO_CLIPBOARD
        .if bCopyFilePathToClipboard
            mov ecx,BST_CHECKED
        .else
            mov ecx,BST_UNCHECKED
        .endif
        Invoke  SendMessage, eax, BM_SETCHECK,ecx,NULL
        ;
        ;*************** Initialize Check Control for bRunWhenWindowsStarts ******************
        ;
        Invoke GetDlgItem,hWnd,IDC_CHK_WIN_STARTUP_RUN
        .if bRunWhenWindowsStarts
            mov ecx,BST_CHECKED
        .else
            mov ecx,BST_UNCHECKED
        .endif
        Invoke  SendMessage, eax, BM_SETCHECK,ecx,NULL
        ;
        ;*************** Initialize Check Control for bHideApplicationWindowWhenOpening ******************
        ;
        Invoke GetDlgItem,hWnd,IDC_CHK_HIDE_WIN_WHEN_OPENING
        .if bHideApplicationWindowWhenOpening
            mov ecx,BST_CHECKED
        .else
            mov ecx,BST_UNCHECKED
        .endif
        Invoke  SendMessage, eax, BM_SETCHECK,ecx,NULL
        ;
        ;*************** Initialize Check Control for bAudibleNotificationAfterCapture ******************
        ;
        Invoke GetDlgItem,hWnd,IDC_CHK_PLAY_AUDIBLE_SHOT
        .if bAudibleNotificationAfterCapture
            mov ecx,BST_CHECKED
        .else
            mov ecx,BST_UNCHECKED
        .endif
        Invoke  SendMessage, eax, BM_SETCHECK,ecx,NULL
        ;
        ;*************** Initialize Check Control for bOpenLastCaptureLocationPath ******************
        ;
        Invoke GetDlgItem,hWnd,IDC_CHK_OPEN_LAST_CAPTURE_PATH
        .if bOpenLastCaptureLocationPath
            mov ecx,BST_CHECKED
        .else
            mov ecx,BST_UNCHECKED
        .endif
        Invoke  SendMessage, eax, BM_SETCHECK,ecx,NULL
        ;
        ;*************** Initialize Check Control for bShowANotificationIcon ******************
        ;
        mov bIsShowANotificationIconActive,FALSE
        Invoke GetDlgItem,hWnd,IDC_CHK_SHOW_NOTIFICATION_ICON
        .if bShowANotificationIcon
            mov ecx,BST_CHECKED
        .else
            mov ecx,BST_UNCHECKED
        .endif
        Invoke  SendMessage, eax, BM_SETCHECK,ecx,NULL
        ;
        ;*************** Initialize Check Control for bShowANotificationIcon ******************
        ;
        Invoke GetDlgItem,hWnd,IDC_CHK_DISPLAY_NOTIFICATION_CAPTURE
        .if bDisplayANotificationCapture
            mov ecx,BST_CHECKED
        .else
            mov ecx,BST_UNCHECKED
        .endif
        Invoke  SendMessage, eax, BM_SETCHECK,ecx,NULL
        ;
        ;*************** Initialize Check Control for bCopyScreenShotToClipboard ******************
        ;
        Invoke GetDlgItem,hWnd,IDC_CHK_COPY_SCREENSHOT_CLIPBOARD
        .if bCopyScreenShotToClipboard
            mov ecx,BST_CHECKED
        .else
            mov ecx,BST_UNCHECKED
        .endif
        Invoke  SendMessage, eax, BM_SETCHECK,ecx,NULL
        ;
        ;*************** Fills up ImageQuality ******************
        ;
        ;*************** Fills up ImageQuality ******************
        ;
        movzx eax,bImageQuality
        Invoke SetDlgItemInt,hWnd,IDC_EDITBOX_IMAGE_QUALITY,eax,FALSE
        ;
        Invoke  GetDlgItem, hWnd, IDC_SPINBOX_IMAGE_QUALITY
        Invoke  SendMessage, eax, UDM_SETRANGE, 0, 00010064H   ;wordLower,wordUpper 1,100
        ;
        ;*************** Fills up ImageZoom ******************
        ;
        movzx eax,bImageZoom
        Invoke SetDlgItemInt,hWnd,IDC_EDITBOX_IMAGE_ZOOM,eax,FALSE
        ;
        Invoke  GetDlgItem, hWnd, IDC_SPINBOX_IMAGE_ZOOM
        Invoke  SendMessage, eax, UDM_SETRANGE, 0, 000100FAH   ;wordLower,wordUpper 1,250
        ;
        ;*************** Fills up Counter ******************
        ;
        mov eax,dSnapShotCounter
        Invoke SetDlgItemInt,hWnd,IDC_EDITBOX_COUNTER,eax,FALSE
        ;
        Invoke  GetDlgItem,hWnd,IDC_SPINBOX_COUNTER
        Invoke  SendMessage,eax,UDM_SETRANGE, 0, 00007FFFH   ;wordLower,wordUpper 0,99
        ;
        ;
        ;Invoke GetDlgItem, hWnd,IDC_BTN_AUTOSHOT
        ;mov hButtonTimer,eax
        ;
        ;
        ;*************** Create PopUpMenu *******************
        ;
        Invoke CreatePopupMenu        ; Create a PopupMenu structure
        mov hMacroPopupMenu, eax           ; and save it's pointer
        ;lazy code below to add items to the menu used when right clicking
        ;
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_01, addr szCOMPUTERNAME
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_02, addr szCOUNTER
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_03, addr szDATE
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_04, addr szDAY
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_05, addr szDRIVE
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_06, addr szERRORCODE
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_07, addr szERRORTEXT
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_08, addr szHOUR
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_09, addr szMINUTE
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_10, addr szMONTH
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_11, addr szPROGRAMDIR
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_12, addr szPROGRAMFILES
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_13, addr szPROGRAMNAME
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_14, addr szSECOND
        ;Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_15, addr szSHARE
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_16, addr szSYSTEMDIR
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_17, addr szSYSTEMDRIVE
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_18, addr szTEMPDIR
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_19, addr szTIME
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_20, addr szUSERNAME
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_21, addr szWEEKDAY
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_22, addr szWINDIR
        Invoke AppendMenu, hMacroPopupMenu, MF_STRING,IDM_23, addr szYEAR
        ;
        ;*************** SubClass EditControl *******************
        ;
        Invoke GetDlgItem, hWnd,IDC_EDITBOX_ARCHIVE_PATH
        mov ecx,0   ;First Element
        mov ActiveIDC[ecx * 8].hIDC,eax
        Invoke SetWindowLong,eax,GWL_WNDPROC,addr SubWndProc
        mov ecx,0
        mov ActiveIDC[ecx * 8].hProc,eax
        ;
        Invoke GetDlgItem, hWnd,IDC_EDITBOX_JPG_COMMENTS
        mov ecx,1   ;Second Element
        mov ActiveIDC[ecx * 8].hIDC,eax
        Invoke SetWindowLong,eax,GWL_WNDPROC,addr SubWndProc
        mov ecx,1
        mov ActiveIDC[ecx * 8].hProc,eax
        ;
        ;
        ;*************** HotKeys *******************
        ;
        szText szActivity_HotKeyRegiterSucesss,"-Hotkey registered sucess: "
        szText szActivity_HotKeyRegiterFail,"-Hotkey registered fail: "
        mov bIsUnhideApplicationWindowHotKeyActive,FALSE
        Invoke  RegisterHotKey,hWnd,HOTKEY_SHIFT_CTRL_ALT_PRINTSCREEN,MOD_NOREPEAT+MOD_SHIFT+MOD_CONTROL+MOD_ALT,VK_SNAPSHOT
        .if eax
            Invoke SetDlgItemText,hWnd,IDC_EDITBOX_HOTKEYS_SHOW_ME_LEFT,addr szHOTKEY_SHIFT_CTRL_ALT_PRINTSCREEN
            mov bIsUnhideApplicationWindowHotKeyActive,TRUE
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterSucesss,addr szHOTKEY_SHIFT_CTRL_ALT_PRINTSCREEN,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .else
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterFail,addr szHOTKEY_SHIFT_CTRL_ALT_PRINTSCREEN,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .endif
        ;
        mov al,'S'
        Invoke  RegisterHotKey,hWnd,HOTKEY_SHIFT_CTRL_ALT_S,MOD_NOREPEAT+MOD_SHIFT+MOD_CONTROL+MOD_ALT,al
        .if eax
            Invoke SetDlgItemText,hWnd,IDC_EDITBOX_HOTKEYS_SHOW_ME_LEFT,addr szHOTKEY_SHIFT_CTRL_ALT_S
            mov bIsUnhideApplicationWindowHotKeyActive,TRUE
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterSucesss,addr szHOTKEY_SHIFT_CTRL_ALT_S,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .else
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterFail,addr szHOTKEY_SHIFT_CTRL_ALT_S,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .endif
        ;
        Invoke  RegisterHotKey, hWnd,HOTKEY_PRINTSCREEN,MOD_NOREPEAT,VK_SNAPSHOT ; MOD_ALT+MOD_CONTROL
        .if eax
            Invoke SetDlgItemText,hWnd,IDC_EDITBOX_HOTKEYS_CAPTURE_TARGET,addr szHOTKEY_PRINTSCREEN
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterSucesss,addr szHOTKEY_PRINTSCREEN,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .else
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterFail,addr szHOTKEY_PRINTSCREEN,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .endif
        ;
        Invoke  RegisterHotKey, hWnd,HOTKEY_CTRL_PRINTSCREEN,MOD_NOREPEAT+MOD_CONTROL,VK_SNAPSHOT ; MOD_ALT+MOD_CONTROL
        .if eax
            Invoke SetDlgItemText,hWnd,IDC_EDITBOX_HOTKEYS_FOR_EDIT_TARGET,addr szHOTKEY_CTRL_PRINTSCREEN
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterSucesss,addr szHOTKEY_CTRL_PRINTSCREEN,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .else
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterFail,addr szHOTKEY_CTRL_PRINTSCREEN,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .endif
        ;
        Invoke  RegisterHotKey, hWnd,HOTKEY_ALT_PRINTSCREEN,MOD_NOREPEAT+MOD_ALT,VK_SNAPSHOT ; MOD_ALT+MOD_CONTROL
        .if eax
            Invoke SetDlgItemText,hWnd,IDC_EDITBOX_HOTKEYS_CAPTURE_TOPWIN,addr szHOTKEY_ALT_PRINTSCREEN
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterSucesss,addr szHOTKEY_ALT_PRINTSCREEN,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .else
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterFail,addr szHOTKEY_ALT_PRINTSCREEN,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .endif
        ;
        Invoke  RegisterHotKey,hWnd,HOTKEY_CTRL_ALT_PRINTSCREEN,MOD_NOREPEAT+MOD_CONTROL+MOD_ALT,VK_SNAPSHOT
        .if eax
            Invoke SetDlgItemText,hWnd,IDC_EDITBOX_HOTKEYS_FOR_EDIT_TOPWIN,addr szHOTKEY_CTRL_ALT_PRINTSCREEN
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterSucesss,addr szHOTKEY_CTRL_ALT_PRINTSCREEN,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .else
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterFail,addr szHOTKEY_CTRL_ALT_PRINTSCREEN,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .endif
        ;
        mov al,'E'
        Invoke  RegisterHotKey,hWnd,HOTKEY_SHIFT_CTRL_E,MOD_NOREPEAT+MOD_SHIFT+MOD_CONTROL,al
        .if eax
            Invoke SetDlgItemText,hWnd,IDC_EDITBOX_HOTKEYS_EMAIL_TARGET,addr szHOTKEY_SHIFT_CTRL_E
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterSucesss,addr szHOTKEY_SHIFT_CTRL_E,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .else
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterFail,addr szHOTKEY_SHIFT_CTRL_E,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .endif
        ;
        mov al,'E'
        Invoke  RegisterHotKey, hWnd,HOTKEY_SHIFT_CTRL_ALT_E,MOD_NOREPEAT+MOD_SHIFT+MOD_CONTROL+MOD_ALT,al
        .if eax
            Invoke SetDlgItemText,hWnd,IDC_EDITBOX_HOTKEYS_EMAIL_TOPWIN,addr szHOTKEY_SHIFT_CTRL_ALT_E
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterSucesss,addr szHOTKEY_SHIFT_CTRL_ALT_E,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .else
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            Invoke UpdateActivityLog,addr szActivity_HotKeyRegiterFail,addr szHOTKEY_SHIFT_CTRL_ALT_E,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .endif
        ;
        ;*************** About History *******************
        Invoke SetDlgItemText,hWnd,IDC_EDITBOX_ABOUT_LOG,addr szApplicationHistory
        ;
        ;*************** MouseHook *******************
        .if bCaptureMode==3
            .if bHookRunning==FALSE
                invoke InstallHook,hWinMain,WM_MOUSEHOOK,NULL
                .if eax==NULL
                    mov bHookRunning,FALSE
                .else
                    mov bHookRunning,TRUE
                .endif
            .endif
        .endif
        ;
    ret
InitDialogWinMain               endp
align 4
ShowMainWindow                  proc bShowWindow:BYTE,dChangeWinMainExStyle:DWORD ;bShowWindow Hide=FALSE/show=TRUE,dChangeWinMainExStyle NULL/NewSlyle
                                ;
        .if dChangeWinMainExStyle==NULL
            .if bShowWindow
                Invoke SetWindowPos,hWinMain,HWND_TOP,0,0,0,0,SWP_SHOWWINDOW OR SWP_NOREPOSITION OR SWP_NOSENDCHANGING OR SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOZORDER
            .else
                Invoke SetWindowPos,hWinMain,HWND_BOTTOM,0,0,0,0,SWP_HIDEWINDOW OR SWP_NOREPOSITION OR SWP_NOSENDCHANGING OR SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOZORDER OR SWP_NOACTIVATE
            .endif
        .else
            Invoke SetWindowLong,hWinMain,GWL_EXSTYLE,dChangeWinMainExStyle
            ;SWP_ASYNCWINDOWPOS: If the calling thread and the thread that owns the window are attached to different input queues, the system posts the request to the thread
                               ;that owns the window. This prevents the calling thread from blocking its execution while other threads process the request.
            ;SWP_DEFERERASE: Prevents generation of the WM_SYNCPAINT message.
            ;SWP_DRAWFRAME: Draws a frame (defined in the window's class description) around the window.
            ;SWP_FRAMECHANGED: Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to the window, even if the window's size
                                ; is not being changed. If this flag is not specified, WM_NCCALCSIZE is sent only when the window's size is being changed.
            ;SWP_NOACTIVATE: Does not activate the window. If this flag is not set, the window is activated and moved to the top of either the topmost or non-topmost group
                                ;(depending on the setting of the hWndInsertAfter parameter).
            ;SWP_NOCOPYBITS: Discards the entire contents of the client area. If this flag is not specified, the valid contents of the client area are saved
                                ;and copied back into the client area after the window is sized or repositioned.
            ;SWP_NOREDRAW: Does not redraw changes. If this flag is set, no repainting of any kind occurs. This applies to the client area, the nonclient area
                                ;(including the title bar and scroll bars), and any part of the parent window uncovered as a result of the window being moved.
                                ;When this flag is set, the application must explicitly invalidate or redraw any parts of the window and parent window that need redrawing.
            ;SWP_NOSENDCHANGING: Prevents the window from receiving the WM_WINDOWPOSCHANGING message.
            ;SWP_NOMOVE,SWP_NOOWNERZORDER,SWP_NOREPOSITION,SWP_NOSIZE,SWP_NOZORDER
            ;SWP_SHOWWINDOW,SWP_HIDEWINDOW
            .if bShowWindow
                ;HWND_BOTTOM, HWND_NOTOPMOST, HWND_TOP,HWND_TOPMOST
                Invoke SetWindowPos,hWinMain,HWND_TOP,0,0,0,0,SWP_SHOWWINDOW OR SWP_FRAMECHANGED OR SWP_NOREPOSITION OR SWP_NOSENDCHANGING OR SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOZORDER
            .else
                ;HWND_BOTTOM, HWND_NOTOPMOST, HWND_TOP,HWND_TOPMOST
                ;Minimiza izq abajo
                ;Invoke SetWindowPos,hWinMain,HWND_BOTTOM,0,0,0,0,SWP_HIDEWINDOW OR SWP_FRAMECHANGED OR SWP_NOREPOSITION OR SWP_NOSENDCHANGING OR SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOZORDER OR SWP_NOACTIVATE
                Invoke SetWindowPos,hWinMain,HWND_BOTTOM,0,0,0,0,SWP_HIDEWINDOW OR SWP_FRAMECHANGED OR SWP_NOREPOSITION OR SWP_NOSENDCHANGING OR SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOZORDER OR SWP_NOACTIVATE
            .endif
        .endif
    ret
ShowMainWindow                  endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
;                       GetThemeFonts
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
GetThemeFonts                   proc
                                LOCAL cHeight:DWORD ;The height, in logical units, of the font's character cell or character.
                                LOCAL cWidth:DWORD ;The average width, in logical units, of characters in the requested font.
                                LOCAL cEscapement:DWORD ;The angle, in tenths of degrees, between the escapement vector and the x-axis of the device. The escapement vector is parallel to the base line of a row of text.
                                LOCAL cOrientation:DWORD ;The angle, in tenths of degrees, between each character's base line and the x-axis of the device.
                                LOCAL cWeight:DWORD ;The weight of the font in the range 0 through 1000. For example, 400 is normal and 700 is bold. If this value is zero, a default weight is used.
                                LOCAL bItalic:DWORD
                                LOCAL bUnderline:DWORD
                                LOCAL bStrikeOut:DWORD
                                LOCAL iQuality:DWORD
        ;Create Theme Mode Fonts and Colors.
        Invoke GetSysColor,COLOR_BTNFACE
        mov LightMode.dColorTransparent,eax
        mov COLOR_NONE,eax
        Invoke CreateSolidBrush,0078D4H
        mov LightMode.hBrushAccent,eax
        mov DarkMode.hBrushAccent,eax
        ;
        xor eax,eax
        mov cHeight,10
        mov cWidth,eax
        mov cEscapement,eax
        mov cOrientation,eax
        mov cWeight,FW_REGULAR
        mov bItalic,eax
        mov bUnderline,eax
        mov bStrikeOut,eax
        mov iQuality,ANTIALIASED_QUALITY ;ANTIALIASED_QUALITY,CLEARTYPE_QUALITY,DEFAULT_QUALITY,DRAFT_QUALITY,NONANTIALIASED_QUALITY,PROOF_QUALITY

        Invoke  CreateFont,cHeight,cWidth,cEscapement,cOrientation,cWeight,bItalic,bUnderline,bStrikeOut,ANSI_CHARSET,OUT_DEFAULT_PRECIS, \
                   CLIP_DEFAULT_PRECIS,iQuality,FF_DONTCARE, addr szSystemFontName
                   ;When you no longer need the font, call the DeleteObject function to delete it.
                   ;If the function succeeds, the return value is a handle to a logical font. If the function fails, the return value is NULL.
        mov LightMode.hRegularFont,eax
        ;
        mov cHeight,38
        mov cWeight,100 ; FW_THIN=100,FW_ULTRALIGHT=200,FW_LIGHT=300, FW_REGULAR=400, FW_MEDIUM=500, FW_BOLD=700
        mov iQuality,ANTIALIASED_QUALITY ;ANTIALIASED_QUALITY,CLEARTYPE_QUALITY,DEFAULT_QUALITY,DRAFT_QUALITY,NONANTIALIASED_QUALITY,PROOF_QUALITY
        Invoke  CreateFont,cHeight,cWidth,cEscapement,cOrientation,cWeight,bItalic,bUnderline,bStrikeOut,ANSI_CHARSET,OUT_DEFAULT_PRECIS, \
                CLIP_DEFAULT_PRECIS,iQuality,FF_DONTCARE, addr szSystemFontName
                   ;When you no longer need the font, call the DeleteObject function to delete it.
                   ;If the function succeeds, the return value is a handle to a logical font. If the function fails, the return value is NULL.
        mov LightMode.hTitleFont,eax
        ;
        mov cHeight,28
        mov cWeight,100 ; FW_THIN=100,FW_ULTRALIGHT=200,FW_LIGHT=300, FW_REGULAR=400, FW_MEDIUM=500, FW_BOLD=700
        mov iQuality,ANTIALIASED_QUALITY ;ANTIALIASED_QUALITY,CLEARTYPE_QUALITY,DEFAULT_QUALITY,DRAFT_QUALITY,NONANTIALIASED_QUALITY,PROOF_QUALITY
        Invoke  CreateFont,cHeight,cWidth,cEscapement,cOrientation,cWeight,bItalic,bUnderline,bStrikeOut,ANSI_CHARSET,OUT_DEFAULT_PRECIS, \
                CLIP_DEFAULT_PRECIS,iQuality,FF_DONTCARE, addr szSystemFontName
                   ;When you no longer need the font, call the DeleteObject function to delete it.
                   ;If the function succeeds, the return value is a handle to a logical font. If the function fails, the return value is NULL.
        mov LightMode.hSubTitleFont,eax

        ret
GetThemeFonts                   endp
;
align 4
SetTheme                        proc hWnd:DWORD,pFonts:DWORD
                                ;
                                LOCAL Rct       :RECT
                                LOCAL hDummy    :DWORD
            ;*************** Apply Theme ******************
            ;
            ;xor eax,eax
            ;mov Rct.left,eax
            ;mov Rct.top,eax
            ;mov Rct.right,240
            ;mov Rct.bottom,60
            ;Invoke CreateFillBmp,FILL_NOISE,addr Rct,30,15,0,0,addr hBmpNoiseNormal,addr hBmpNoiseHover
            ;
            ;*************** Colors for all Buttons ******************
            ;
            mov hIconOnOffAscent,   FUNC(LoadIcon,hInstance,502)
            mov hIconOnOffUnpush,   FUNC(LoadIcon,hInstance,503)
            mov hIconCameraAscent,  FUNC(LoadIcon,hInstance,504)
            mov hIconCameraUnpush,  FUNC(LoadIcon,hInstance,505)
            mov hIconFileAscent,    FUNC(LoadIcon,hInstance,506)
            mov hIconFileUnpush,    FUNC(LoadIcon,hInstance,507)
            mov hIconSettingsAscent,FUNC(LoadIcon,hInstance,508)
            mov hIconSettingsUnpush,FUNC(LoadIcon,hInstance,509)
            mov hIconHotkeysAscent, FUNC(LoadIcon,hInstance,510)
            mov hIconHotkeysUnpush, FUNC(LoadIcon,hInstance,511)
            mov hIconWarningAscent, FUNC(LoadIcon,hInstance,512)
            mov hIconWarningUnpush, FUNC(LoadIcon,hInstance,513)
            mov hIconAboutAscent,   FUNC(LoadIcon,hInstance,514)
            mov hIconAboutUnpush,   FUNC(LoadIcon,hInstance,515)
            mov hIconExitAscent,    FUNC(LoadIcon,hInstance,516)
            mov hIconExitUnpush,    FUNC(LoadIcon,hInstance,517)
            mov hIconSaveAscent,    FUNC(LoadIcon,hInstance,518)
            mov hIconSaveUnpush,    FUNC(LoadIcon,hInstance,519)
            mov hIconAutoshotAscent,FUNC(LoadIcon,hInstance,520)
            mov hIconAutoshotUnpush,FUNC(LoadIcon,hInstance,521)
            mov hIconToMailAscent,  FUNC(LoadIcon,hInstance,522)
            mov hIconToMailUnpush,  FUNC(LoadIcon,hInstance,523)
            mov hIconToClipAscent,  FUNC(LoadIcon,hInstance,524)
            mov hIconToClipUnpush,  FUNC(LoadIcon,hInstance,525)
            mov hIconHideAscent,    FUNC(LoadIcon,hInstance,526)
            mov hIconHideUnpush,    FUNC(LoadIcon,hInstance,527)
            mov hIconToPathAscent,  FUNC(LoadIcon,hInstance,528)
            mov hIconToPathUnpush,  FUNC(LoadIcon,hInstance,529)
            ;
            Invoke FillBuffer,addr BtnStruct, sizeof BtnStruct,0
            mov BtnStruct.hCursor_hover, NULL
            mov BtnStruct.hover_clr, $RGB (0,120,212)
            mov BtnStruct.push_clr, $RGB (0,120,212)
            mov BtnStruct.normal_clr, $RGB (0,0,0)
            ;BUTTONS PROPS: DRAW_BORDER,FLAT_BTN,DRAW_FOCUS,TRANSPARENT_BKGND,RGN_BUTTON,BALLOON_TT,ICON_LEFT,ICON_RIGHT,ICON_TOP
            ;TRANSPARENT_BKGND = Invoke FillRect,hBtnDC,addr RctBtn,COLOR_BTNFACE+1
            ;
            ;STYLE_DEFAULT,STYLE_XNET,STYLE_OFFICE_XP,STYLE_OFFICE_2000
            ;mov BtnStruct.btn_prop, STYLE_XNET+FLAT_BTN+DRAW_BORDER+DRAW_FOCUS+TRANSPARENT_BKGND
            ;mov BtnStruct.btn_prop, STYLE_DEFAULT+FLAT_BTN+TRANSPARENT_BKGND
            ;mov BtnStruct.btn_prop, STYLE_OFFICE_XP+FLAT_BTN+TRANSPARENT_BKGND
            ;mov BtnStruct.btn_prop, STYLE_OFFICE_2000+FLAT_BTN+TRANSPARENT_BKGND
            ;mov BtnStruct.btn_prop, FLAT_BTN+TRANSPARENT_BKGND+RGN_BUTTON ;Sólo texto
            ;mov BtnStruct.btn_prop, FLAT_BTN+STYLE_XNET+TRANSPARENT_BKGND
            ;mov BtnStruct.btn_prop, FLAT_BTN+STYLE_XNET
            ;mov BtnStruct.btn_prop, FLAT_BTN+DRAW_BORDER+STYLE_XNET+DRAW_FOCUS+TRANSPARENT_BKGND+RGN_BUTTON
            ;mov BtnStruct.btn_prop, FLAT_BTN+DRAW_BORDER+DRAW_FOCUS+TRANSPARENT_BKGND+RGN_BUTTON ;Igual que el fondo
            ;mov BtnStruct.btn_prop, FLAT_BTN+DRAW_BORDER+DRAW_FOCUS+TRANSPARENT_BKGND
            ;
            ;*************** LIST OPTIONS Buttons ******************
            ;BUTTONS PROPS: DRAW_BORDER,FLAT_BTN,DRAW_FOCUS,TRANSPARENT_BKGND,RGN_BUTTON,BALLOON_TT,ICON_LEFT,ICON_RIGHT,ICON_TOP
            ;TODO ;Invoke Button_SetTextMargin,hDummy,addr Rct ;Comctl32.dll
            mov BtnStruct.btn_prop,FLAT_BTN+TRANSPARENT_BKGND+RGN_BUTTON+ICON_LEFT ;+DRAW_FOCUS ICON_LEFT
            ;
            mov eax,hIconCameraUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconCameraAscent
            mov BtnStruct.hIcon_hover,eax
            mov BtnStruct.hIcon_push,eax
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_LIST_ITEM_BUTTON1)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov eax,hIconFileUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconFileAscent
            mov BtnStruct.hIcon_hover,eax
            mov BtnStruct.hIcon_push,eax
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_LIST_ITEM_BUTTON2)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov eax,hIconSettingsUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconSettingsAscent
            mov BtnStruct.hIcon_hover,eax
            mov BtnStruct.hIcon_push,eax
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_LIST_ITEM_BUTTON3)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov eax,hIconHotkeysUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconHotkeysAscent
            mov BtnStruct.hIcon_hover,eax
            mov BtnStruct.hIcon_push,eax
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_LIST_ITEM_BUTTON4)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov eax,hIconWarningUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconWarningAscent
            mov BtnStruct.hIcon_hover,eax
            mov BtnStruct.hIcon_push,eax
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_LIST_ITEM_BUTTON5)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov eax,hIconAboutUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconAboutAscent
            mov BtnStruct.hIcon_hover,eax
            mov BtnStruct.hIcon_push,eax
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_LIST_ITEM_BUTTON6)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov eax,hIconSaveUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconSaveAscent
            mov BtnStruct.hIcon_hover,eax
            mov BtnStruct.hIcon_push,eax
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_LIST_ITEM_BUTTON7)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov eax,hIconExitUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconExitAscent
            mov BtnStruct.hIcon_hover,eax
            mov BtnStruct.hIcon_push,eax
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_LIST_ITEM_BUTTON8)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            ;*************** CAPTURE PAGE Execute Buttons ******************
            ;BUTTONS PROPS: DRAW_BORDER,FLAT_BTN,DRAW_FOCUS,TRANSPARENT_BKGND,RGN_BUTTON,BALLOON_TT,ICON_LEFT,ICON_RIGHT,ICON_TOP
            mov BtnStruct.btn_prop,ICON_TOP ;+DRAW_FOCUS ICON_LEFT ICON_RIGHT ICON_TOP +TRANSPARENT_BKGND
            mov BtnStruct.hover_clr, $RGB (0,120,212)
            mov BtnStruct.push_clr, $RGB (0,120,212)
            ;
            mov eax,hIconToPathUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconToPathAscent
            mov BtnStruct.hIcon_push,eax
            mov BtnStruct.hIcon_hover,eax
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_BTN_SEND_TO_LOCATION)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov eax,hIconToMailUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconToMailAscent
            mov BtnStruct.hIcon_push,eax
            mov BtnStruct.hIcon_hover,eax
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_BTN_SEND_TO_EMAIL)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov eax,hIconToClipUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconToClipAscent
            mov BtnStruct.hIcon_push,eax
            mov BtnStruct.hIcon_hover,eax
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_BTN_SEND_TO_CLIPBOARD)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov eax,hIconHideUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconHideAscent
            mov BtnStruct.hIcon_push,eax
            mov BtnStruct.hIcon_hover,eax
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_BTN_HIDE)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov BtnStruct.btn_prop,FLAT_BTN+RGN_BUTTON+ICON_TOP ;+DRAW_FOCUS ICON_LEFT
            mov eax,hIconAutoshotUnpush
            mov BtnStruct.hIcon_normal,eax
            mov eax,hIconAutoshotAscent
            mov BtnStruct.hIcon_hover,eax
            mov BtnStruct.hIcon_push,eax
            lea eax,szText_StartAutoShot
            mov BtnStruct.lpText_normal,eax
            lea eax,szText_StopAutoShot
            mov BtnStruct.lpText_push,eax
            mov BtnStruct.hover_clr, $RGB (178,21,0)
            ;mov BtnStruct.push_clr, $RGB (178,21,0)
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_BTN_AUTOSHOT)
            Invoke SendMessage,hDummy,BCM_SETTEXTMARGIN,NULL,addr Rct
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            ;*************** Check On/Off Buttons ******************
            mov eax,hIconOnOffUnpush
            mov BtnStruct.hIcon_normal,eax
            mov BtnStruct.hIcon_hover,eax
            mov eax,hIconOnOffAscent
            mov BtnStruct.hIcon_push,eax
            lea eax,szText_normal
            mov BtnStruct.lpText_normal,eax
            lea eax,szText_push
            mov BtnStruct.lpText_push,eax
            mov BtnStruct.hover_clr, $RGB (0,120,212)
            mov BtnStruct.push_clr, $RGB (0,120,212)
            mov BtnStruct.btn_prop,FLAT_BTN+RGN_BUTTON ;+DRAW_FOCUS ICON_LEFT
            ;mov BtnStruct.btn_prop, FLAT_BTN+DRAW_BORDER+DRAW_FOCUS+TRANSPARENT_BKGND+RGN_BUTTON ;Igual que el fondo
            ;
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_CHK_START_AUTOSHOT_AT_INIT)
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_CHK_COPY_PATH_TO_CLIPBOARD)
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_CHK_WIN_STARTUP_RUN)
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_CHK_HIDE_WIN_WHEN_OPENING)
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_CHK_PLAY_AUDIBLE_SHOT)
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_CHK_OPEN_LAST_CAPTURE_PATH)
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_CHK_SHOW_NOTIFICATION_ICON)
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_CHK_DISPLAY_NOTIFICATION_CAPTURE)
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;
            mov hDummy,$invoke(GetDlgItem,hWnd,IDC_CHK_COPY_SCREENSHOT_CLIPBOARD)
            Invoke RedrawButton,hDummy,addr BtnStruct
            ;*************** Title / Subtitle ******************
            ;SendMessage The low-order word of lParam specifies whether the control should be redrawn immediately upon setting the font. If this parameter is TRUE, the control redraws itself.
            Invoke GetDlgItem,hWnd,IDC_PROMPT_TITLE_FILE
            mov ecx,LightMode.hTitleFont
            Invoke  SendMessage,eax, WM_SETFONT, ecx,FALSE
            ;
            Invoke GetDlgItem,hWnd,IDC_PROMPT_SUBTITLE_FILE
            mov ecx,LightMode.hSubTitleFont
            Invoke  SendMessage,eax, WM_SETFONT, ecx,FALSE
            ;
            Invoke GetDlgItem,hWnd,IDC_PROMPT_TITLE_CAPTURE
            mov ecx,LightMode.hTitleFont
            Invoke  SendMessage,eax, WM_SETFONT, ecx,FALSE
            ;
            Invoke GetDlgItem,hWnd,IDC_PROMPT_SUBTITLE_CAPTURE
            mov ecx,LightMode.hSubTitleFont
            Invoke  SendMessage,eax, WM_SETFONT, ecx,FALSE
            ;
            Invoke GetDlgItem,hWnd,IDC_PROMPT_TITLE_APP
            mov ecx,LightMode.hTitleFont
            Invoke  SendMessage,eax, WM_SETFONT, ecx,FALSE
            ;
            Invoke GetDlgItem,hWnd,IDC_PROMPT_SUBTITLE_APP
            mov ecx,LightMode.hSubTitleFont
            Invoke  SendMessage,eax, WM_SETFONT, ecx,FALSE
            ;
            Invoke GetDlgItem,hWnd,IDC_PROMPT_HOTKEYS_TITLE
            mov ecx,LightMode.hTitleFont
            Invoke  SendMessage,eax, WM_SETFONT, ecx,FALSE
            ;
            Invoke GetDlgItem,hWnd,IDC_PROMPT_HOTKEYS_SUBTITLE
            mov ecx,LightMode.hSubTitleFont
            Invoke  SendMessage,eax, WM_SETFONT, ecx,FALSE
            ;
            Invoke GetDlgItem,hWnd,IDC_PROMPT_ACTIVITY_TITLE
            mov ecx,LightMode.hTitleFont
            Invoke  SendMessage,eax, WM_SETFONT, ecx,FALSE
            ;
            Invoke GetDlgItem,hWnd,IDC_PROMPT_ACTIVITY_SUBTITLE
            mov ecx,LightMode.hSubTitleFont
            Invoke  SendMessage,eax, WM_SETFONT, ecx,FALSE
            ;
            Invoke GetDlgItem,hWnd,IDC_PROMPT_ABOUT_TITLE
            mov ecx,LightMode.hTitleFont
            Invoke  SendMessage,eax, WM_SETFONT, ecx,FALSE
            ;
            Invoke GetDlgItem,hWnd,IDC_PROMPT_ABOUT_SUBTITLE
            mov ecx,LightMode.hSubTitleFont
            Invoke  SendMessage,eax, WM_SETFONT, ecx,FALSE
    ret
SetTheme                        endp
;
align 4
SelectPage                      proc dPageID:DWORD ;Select Active Page: Shows controls from page ID and hides others
                                LOCAL hTab:DWORD
        mov eax,dPageID
        .if dLastSelectedPage == eax
            jmp Exit_SelectPage
        .else
            mov dLastSelectedPage,eax
        .endif
        mov bSelectingPage,TRUE
        ;
;        Invoke GetDlgItem,hWinMain,IDC_SHEET_TAB_CONTROL_COVER
;        mov hTab,eax
;        Invoke ShowWindow,hTab,SW_SHOW
;        Invoke SetWindowPos,hTab,HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE OR SWP_NOSIZE OR SWP_SHOWWINDOW ;To mov Cover to top, all other controsl must have "WS_CLIPSIBLINGS"
        ;Hide all Pages.
        mov ecx,IDC_PAGE01
        @@:
            push ecx
            Invoke ShowPage,ecx,SW_HIDE
            pop ecx
            add ecx,1
            cmp ecx,IDC_LASTPAGE
            jle @B
        ;
        ;Show Selected Page.
        Invoke ShowPage,dPageID,SW_SHOW
;        Invoke SetWindowPos,hTab,HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE OR SWP_NOSIZE OR SWP_HIDEWINDOW
;        Invoke ShowPage,dPageID,SW_SHOW
        ;
        Exit_SelectPage:
        mov bSelectingPage,FALSE
    ret
SelectPage                      endp
;
ShowPage                        proc dPageID:DWORD,nCmdShow:DWORD

        ;
        .if dPageID == IDC_PAGE01
            mov ecx,IDC_PROMPT_TITLE_CAPTURE
            @@:
                push ecx
                Invoke GetDlgItem,hWinMain,ecx
                Invoke ShowWindow,eax,nCmdShow
                pop ecx
                add ecx,1
                cmp ecx,IDC_BTN_HIDE
                jle @B
            Invoke GetDlgItem,hWinMain,IDC_BTN_SAVE
            Invoke ShowWindow,eax,SW_HIDE ;Disabled in SelectPage
        .elseif dPageID == IDC_PAGE02
            mov ecx,IDC_PROMPT_TITLE_FILE
            @@:
                push ecx
                Invoke GetDlgItem,hWinMain,ecx
                Invoke ShowWindow,eax,nCmdShow
                pop ecx
                add ecx,1
                cmp ecx,IDC_EDITBOX_JPG_COMMENTS
                jle @B
        .elseif dPageID == IDC_PAGE03
            mov ecx,IDC_PROMPT_TITLE_APP
            @@:
                push ecx
                Invoke GetDlgItem,hWinMain,ecx
                Invoke ShowWindow,eax,nCmdShow
                pop ecx
                add ecx,1
                cmp ecx,IDC_CHK_DISPLAY_NOTIFICATION_CAPTURE
                jle @B
        .elseif dPageID == IDC_PAGE04
            mov ecx,IDC_PROMPT_HOTKEYS_TITLE
            @@:
                push ecx
                Invoke GetDlgItem,hWinMain,ecx
                Invoke ShowWindow,eax,nCmdShow
                pop ecx
                add ecx,1
                cmp ecx,IDC_EDITBOX_HOTKEYS_EMAIL_TOPWIN
                jle @B
        .elseif dPageID == IDC_PAGE05
            mov ecx,IDC_PROMPT_ACTIVITY_TITLE
            @@:
                push ecx
                Invoke GetDlgItem,hWinMain,ecx
                Invoke ShowWindow,eax,nCmdShow
                pop ecx
                add ecx,1
                cmp ecx,IDC_EDITBOX_ACTIVITY_LOG
                jle @B
        .elseif dPageID == IDC_PAGE06
            mov ecx,IDC_PROMPT_ABOUT_TITLE
            @@:
                push ecx
                Invoke GetDlgItem,hWinMain,ecx
                Invoke ShowWindow,eax,nCmdShow
                pop ecx
                add ecx,1
                cmp ecx,IDC_EDITBOX_ABOUT_LOG
                jle @B
        ;.elseif dPageID = IDC_PAGE07
        ;.elseif dPageID = IDC_PAGE08
        ;.elseif dPageID = IDC_PAGE09
        .endif
        ;
    ret
ShowPage                        endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;TakeScreenShot and AutoShotLoop
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
TakeScreenShot                  proc bLocalCaptureWindow:BYTE,dSendArchiveTo:BYTE
                            ;LocalData
                            Local Param_aawin2jpg_WindowsTarget:DWORD
                            LOCAL szSolvedJpgPath[MAX_PATH]:BYTE
                            LOCAL szShellExecuteJpgPath[MAX_PATH+MAX_PATH]:BYTE
                            LOCAL szSolvedJpgComments[MAX_PATH]:BYTE
                            LOCAL dBytesWritten:DWORD
                            LOCAL pBlobData:DWORD
                            LOCAL hClipData:DWORD
                            LOCAL pClipData:DWORD
                            LOCAL dClipDataSize:DWORD
                            LOCAL dCaptureWindow:DWORD
                            LOCAL bGray:BYTE
                            LOCAL hDesktopWinHnd:DWORD
                            LOCAL hDC:DWORD
                            LOCAL hWinActive:DWORD
                            LOCAL hMonitor
                            LOCAL lpmi:MONITORINFO
                            LOCAL dWidth:DWORD
                            LOCAL dHeight:DWORD
                            LOCAL drop:DROPFILES
                            ;
                          ;
                          ;

                            ;LOCAL shop:SHFILEOPSTRUCT
                            ;LOCAL dAbort:DWORD
;                            typedef struct _SHFILEOPSTRUCT {
;                              HWND         hwnd;
;                              UINT         wFunc;
;                              PCZZTSTR     pFrom;
;                              PCZZTSTR     pTo;
;                              FILEOP_FLAGS fFlags;
;                              BOOL         fAnyOperationsAborted;
;                              LPVOID       hNameMappings;
;                              PCTSTR       lpszProgressTitle;
;                            } SHFILEOPSTRUCT, *LPSHFILEOPSTRUCT;
        ;
        Invoke IniManager,hWinMain,1   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
        ;
        .if bCaptureMode==RUNTIME_CAPTURE_MODE_CLICKS
            .if bHookRunning==FALSE
                Invoke InstallHook,hWinMain,WM_MOUSEHOOK,NULL
                .if eax==NULL
                    mov bHookRunning,FALSE
                .else
                    mov bHookRunning,TRUE
                .endif
            .endif
        .else
            .if bHookRunning==TRUE
                Invoke UninstallHook
                mov bHookRunning,FALSE
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke UpdateActivityLog,NULL,NULL,0,0
        szText szActivity_StartScreenshot_Title,9,9,9,"*** New call for screen capture begins ***"
        Invoke UpdateActivityLog,addr szActivity_StartScreenshot_Title,NULL,0,0
        szText szActivity_StartScreenShot_UnparsedPath,"Unparsed Path: "
        Invoke UpdateActivityLog,addr szActivity_StartScreenShot_UnparsedPath,addr szArchivePath,2,0
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        ;

        Invoke FillBuffer,addr szSolvedJpgPath, sizeof szSolvedJpgPath,NULL
        ;********* Thread is Saving process
        Invoke SolveMacros,addr szArchivePath,addr szSolvedJpgPath
        Invoke SolveMacros,addr szJpgComments,addr szSolvedJpgComments
        Invoke lstrcpy, addr szImageFilePath,addr szSolvedJpgPath
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        szText szActivity_StartScreenShot_ParsedFile,"Parsed File: "
        Invoke UpdateActivityLog,addr szActivity_StartScreenShot_ParsedFile,addr szSolvedJpgPath,1,0
        szText szActivity_StartScreenShot_ParsedJpgComments,"Jpg Comments: "
        Invoke UpdateActivityLog,addr szActivity_StartScreenShot_ParsedJpgComments,addr szSolvedJpgComments,1,0
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        ;
        mov dBytesWritten,0
        mov al,bImageMode
        dec al
        mov bGray,al
;        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
;        szText szActivity_StartScreenShot_ImageModeTitle,"Image Mode: "
;        szText szActivity_StartScreenShot_ImageModeColor,"Color"
;        szText szActivity_StartScreenShot_ImageModeBlackAndWhite,"Black and White"
;        .if bGray
;            Invoke UpdateActivityLog,addr szActivity_StartScreenShot_ImageModeTitle,addr szActivity_StartScreenShot_ImageModeBlackAndWhite,1,0
;        .else
;            Invoke UpdateActivityLog,addr szActivity_StartScreenShot_ImageModeTitle,addr szActivity_StartScreenShot_ImageModeColor,1,0
;        .endif
;        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬

        .if (bLocalCaptureWindow == RUNTIME_TARGET_TYPE_ACTIVE_WINDOWS) || (bLocalCaptureWindow == RUNTIME_TARGET_TYPE_DESKTOP)  ;0= Active Monitor, 1=Active Window, 2=Virtual Desktop, 3=Region

        ;######################## MAKE WINDOW SNAPSHOT ###################
            .if bLocalCaptureWindow == RUNTIME_TARGET_TYPE_ACTIVE_WINDOWS ;-1=Active Window, 0=Virtual Desktop
                mov Param_aawin2jpg_WindowsTarget,-1
            .else
                mov Param_aawin2jpg_WindowsTarget,0
            .endif
            ;aawin2jpg,lpSize,WinHnd,Address,AddrType,Quality(1-100),Zoom,GrayScale,Comments
            ;db 42 dup('Disable',0,0,0,0,0,0,0,'Default Arrow',0,'Big Red Arrow',0);14
            mov cl,101
            sub cl,bImageQuality
            .if bCaptureCursor == RUNTIME_CAPTURE_CURSOR_DISABLE ;,RUNTIME_CAPTURE_CURSOR_SYSTEM_ARROW,RUNTIME_CAPTURE_CURSOR_RED_ARROW
                Invoke aawin2jpg,addr dBytesWritten,Param_aawin2jpg_WindowsTarget,addr szSolvedJpgPath,0,cl,bImageZoom,bGray,addr szSolvedJpgComments
            .else
                .if bCaptureCursor == RUNTIME_CAPTURE_CURSOR_SYSTEM_ARROW
                    Invoke aawinwithcursor2jpg,addr dBytesWritten,Param_aawin2jpg_WindowsTarget,addr szSolvedJpgPath,0,cl,bImageZoom,bGray,addr szSolvedJpgComments,hCursor
                .else
                    Invoke aawinwithcursor2jpg,addr dBytesWritten,Param_aawin2jpg_WindowsTarget,addr szSolvedJpgPath,0,cl,bImageZoom,bGray,addr szSolvedJpgComments,hCursorRed
                .endif
            .endif
            jmp AfterCaptureActions
            ;
        .else
        ;######################## MAKE REGION SNAPSHOT ###################
            .if bLocalCaptureWindow == RUNTIME_TARGET_TYPE_REGION
                Invoke IniManager,0,4   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
                mov cl,101
                sub cl,bImageQuality
                .if bCaptureCursor == RUNTIME_CAPTURE_CURSOR_DISABLE ;,RUNTIME_CAPTURE_CURSOR_SYSTEM_ARROW,RUNTIME_CAPTURE_CURSOR_RED_ARROW
                    Invoke aawin2jpg,addr dBytesWritten,addr rcCapture,addr szSolvedJpgPath,0,cl,bImageZoom,bGray,addr szSolvedJpgComments
                .else
                    .if bCaptureCursor == RUNTIME_CAPTURE_CURSOR_SYSTEM_ARROW
                        Invoke aawinwithcursor2jpg,addr dBytesWritten,addr rcCapture,addr szSolvedJpgPath,0,cl,bImageZoom,bGray,addr szSolvedJpgComments,hCursor
                    .else
                        Invoke aawinwithcursor2jpg,addr dBytesWritten,addr rcCapture,addr szSolvedJpgPath,0,cl,bImageZoom,bGray,addr szSolvedJpgComments,hCursorRed
                    .endif
                .endif
                jmp AfterCaptureActions
                ;
            .else ;RUNTIME_TARGET_TYPE_ACTIVE_MONITOR
                Invoke GetForegroundWindow
                .if eax==NULL
                    Invoke Sleep,10 ;The foreground window can be NULL in certain circumstances, such as when a window is losing activation.
                    Invoke GetForegroundWindow
                    .if eax==NULL
                        Invoke GetActiveWindow
                    .endif
                .endif
                mov hWinActive,eax
                Invoke IsWindow,hWinActive
                .if eax==FALSE
                    Invoke GetDesktopWindow
                    mov hWinActive,eax
                .endif
                ;MonitorFromPoint MonitorFromRect
                Invoke MonitorFromWindow,hWinActive,MONITOR_DEFAULTTONEAREST ;MONITOR_DEFAULTTONULL ;MONITOR_DEFAULTTOPRIMARY
                mov hMonitor,eax
                ;MONITORINFO STRUCT
                ;    DWORD cbSize     ;The size of the structure, in bytes.
                ;    RECT  rcMonitor  ;A RECT structure that specifies the display monitor rectangle, expressed in virtual-screen coordinates. Note that if the monitor is not the primary display monitor, some of the rectangle's coordinates may be negative values.
                ;    RECT  rcWork     ;A RECT structure that specifies the work area rectangle of the display monitor, expressed in virtual-screen coordinates. Note that if the monitor is not the primary display monitor, some of the rectangle's coordinates may be negative values.
                ;    DWORD dwFlags    ;A set of flags that represent attributes of the display monitor: MONITORINFOF_PRIMARY
                ;MONITORINFO ENDS
                mov eax,sizeof lpmi
                lea edx,lpmi
                assume edx:PTR MONITORINFO
                    mov [edx].cbSize, sizeof MONITORINFO
                assume edx:nothing
                Invoke GetMonitorInfoA,hMonitor,addr lpmi
                ;
                 mov cl,101
                 sub cl,bImageQuality
                .if bCaptureCursor == RUNTIME_CAPTURE_CURSOR_DISABLE ;,RUNTIME_CAPTURE_CURSOR_SYSTEM_ARROW,RUNTIME_CAPTURE_CURSOR_RED_ARROW
                    Invoke aawin2jpg,addr dBytesWritten,addr lpmi.rcMonitor,addr szSolvedJpgPath,0,cl,bImageZoom,bGray,addr szSolvedJpgComments
                .else
                    .if bCaptureCursor == RUNTIME_CAPTURE_CURSOR_SYSTEM_ARROW
                        Invoke aawinwithcursor2jpg,addr dBytesWritten,addr lpmi.rcMonitor,addr szSolvedJpgPath,0,cl,bImageZoom,bGray,addr szSolvedJpgComments,hCursor
                    .else
                        Invoke aawinwithcursor2jpg,addr dBytesWritten,addr lpmi.rcMonitor,addr szSolvedJpgPath,0,cl,bImageZoom,bGray,addr szSolvedJpgComments,hCursorRed
                    .endif
                .endif
                jmp AfterCaptureActions
                ;
            .endif
        .endif
        AfterCaptureActions:
        ;##########################################################
        ;                     ACTIVITY / ERROR
        ;##########################################################
        .if eax!=0
            jmp ErrorShot
        .else
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            szText szActivity_StartScreenShot_EncoderSucess,"File Encoding: Sucess."
            Invoke UpdateActivityLog,addr szActivity_StartScreenShot_EncoderSucess,NULL,1,0
            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .endif
        ;##########################################################
        ;           UPDATE %COUNTER%
        ;##########################################################
        Invoke IniManager,hWinMain,6   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
        add dSnapShotCounter,1
        Invoke SetDlgItemInt,hWinMain,IDC_EDITBOX_COUNTER,dSnapShotCounter,FALSE
        Invoke IniManager,hWinMain,7   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
        ;##########################################################
        ;
        ;##########################################################
        ;                     PLAY SOUND
        ;##########################################################
        .if bAudibleNotificationAfterCapture
            ;800 WAVE DISCARDABLE "./images/camera.wav"
            Invoke PlaySound,800,hInstance,SND_RESOURCE OR SND_ASYNC OR SND_SENTRY OR SND_SYSTEM ;SND_ASYNC ;OR SND_LOOP
            ;Invoke PlaySound,800,hInstance, SND_RESOURCE OR SND_SYNC OR SND_SENTRY OR SND_SYSTEM ;SND_ASYNC ;OR SND_LOOP
            ;Invoke Sleep,100
            ;Invoke PlaySound,NULL,hInstance,SND_LOOP
            ;Invoke PlaySound,800,hInstance,SND_RESOURCE OR SND_SYSTEM OR SND_ASYNC OR SND_NOSTOP;or SND_SYNC SND_SYSTEM

        .endif
        ;##########################################################
        ;               COPY IMAGE TO CLIPBOARD
        ;##########################################################
        .if bCopyFilePathToClipboard==TRUE || bCopyScreenShotToClipboard==TRUE || dSendArchiveTo==RUNTIME_CAPTURE_SENT_TO_CLIPBOARD || dSendArchiveTo==RUNTIME_CAPTURE_SENT_TO_MAIL
            Invoke OpenClipboard,hWinMain
            .if eax==NULL ;If the function succeeds, the return value is nonzero.
                Invoke aaerrortext,eax
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                szText szActivity_StartScreenShot_ClipboardOpeningError,"Open Clipboard error: "
                Invoke UpdateActivityLog,addr szActivity_StartScreenShot_ClipboardOpeningError,eax,1,0
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            .else
                ;Invoke MessageBox,NULL,addr szSolvedJpgPath,addr szSolvedJpgComments,NULL
                Invoke EmptyClipboard ; always empty the clipboard before writing to it !!!
                ;---------------------------------------------------------------------------
                ;                     DROP: COPIAR EL ARCHIVO
                ;---------------------------------------------------------------------------
                Invoke lstrlen,addr szSolvedJpgPath
                add eax,3
                add eax,sizeof drop
                mov dClipDataSize,eax
                Invoke GlobalAlloc,GMEM_MOVEABLE+GMEM_DDESHARE,dClipDataSize;dBytesWritten ;If the hMem parameter identifies a memory object, the object must have been allocated using the function with the GMEM_MOVEABLE flag.
                ;Invoke GlobalAlloc,GMEM_MOVEABLE,dClipDataSize ;dBytesWritten
                mov hClipData, eax
                .if hClipData
                    Invoke GlobalLock,hClipData
                    mov pClipData,eax
                    Invoke FillBuffer,pClipData,dClipDataSize,NULL
                    mov edx,pClipData
                    assume edx:PTR DROPFILES
                        mov [edx].pFiles,sizeof DROPFILES ;drop
                        ;mov [edx].pt.
                        ;mov [edx].pt.
                        ;mov [edx].fNC,FALSE      ;NonClient Area
                        ;mov [edx],fWide,FALSE    ;Ansi=FAlse, Unicode=True
                    assume edx:nothing
                    mov edx,pClipData
                    add edx, sizeof DROPFILES ;drop ;DROPFILES
                    Invoke lstrcpy,edx,addr szSolvedJpgPath
                    Invoke GlobalUnlock,hClipData ;(The memory must be unlocked before the Clipboard is closed.)
                    or eax,eax
                    jnz @F
                        Invoke SetClipboardData,CF_HDROP,hClipData ;If SetClipboardData succeeds, the system owns the object identified by the hMem parameter.
                        ;If the function succeeds, the return value is the handle to the data. If the function fails, the return value is NULL
                        .if eax==NULL
                            Invoke GetLastError
                            Invoke aaerrortext,eax
                            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                            szText szActivity_StartScreenShot_ClipboardSetImageError,"Error putting Image to Clipboard: "
                            Invoke UpdateActivityLog,addr szActivity_StartScreenShot_ClipboardSetImageError,eax,1,0
                            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                            Invoke GlobalFree,hClipData
                        .else
                            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                            szText szActivity_StartScreenShot_ClipboardSetImageSuccess,"Copy Image to Clipboard: Success."
                            Invoke UpdateActivityLog,addr szActivity_StartScreenShot_ClipboardSetImageSuccess,NULL,1,0
                            ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                        .endif
                    @@:
                .endif
                ;---------------------------------------------------------------------------
                ;                     TEXT: COPIAR PATH DEL ARCHIVO
                ;---------------------------------------------------------------------------
                .if bCopyFilePathToClipboard==TRUE || dSendArchiveTo==RUNTIME_CAPTURE_SENT_TO_CLIPBOARD
                    Invoke lstrlen,addr szSolvedJpgPath
                    add eax,1
                    mov dClipDataSize,eax
                    Invoke GlobalAlloc,GMEM_MOVEABLE+GMEM_DDESHARE,dClipDataSize ;If the hMem parameter identifies a memory object, the object must have been allocated using the function with the GMEM_MOVEABLE flag.
                    mov hClipData, eax
                    .if hClipData
                        Invoke GlobalLock,hClipData
                        mov pClipData,eax
                        Invoke lstrcpy,pClipData,addr szSolvedJpgPath
                        Invoke GlobalUnlock,hClipData ;(The memory must be unlocked before the Clipboard is closed.)
                        or eax,eax
                        jnz @F
                            Invoke SetClipboardData,CF_TEXT,hClipData ;If SetClipboardData succeeds, the system owns the object identified by the hMem parameter.
                            ;If the function succeeds, the return value is the handle to the data. If the function fails, the return value is NULL
                            .if eax==NULL
                                Invoke GetLastError
                                Invoke aaerrortext,eax
                                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                                szText szActivity_StartScreenShot_ClipboardSetPathError,"Error putting file path to Clipboard: "
                                Invoke UpdateActivityLog,addr szActivity_StartScreenShot_ClipboardSetPathError,eax,1,0
                                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                                Invoke GlobalFree,hClipData
                            .else
                                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                                szText szActivity_StartScreenShot_ClipboardSetPathSuccess,"Copy file path to Clipboard: Success."
                                Invoke UpdateActivityLog,addr szActivity_StartScreenShot_ClipboardSetPathSuccess,NULL,1,0
                                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                            .endif
                        @@:
                    .endif
                .endif
                Invoke CloseClipboard
            .endif
            ;
            ;
        .endif
        ;##########################################################
        ;                     DISPLAY NOTIFICATION
        ;##########################################################
        .if bShowANotificationIcon && bDisplayANotificationCapture
            ;********** Copy File Name to Balloon Tip *****
            Invoke lstrlen, addr szSolvedJpgPath
            mov ecx,eax
            lea edx,szSolvedJpgPath
            add eax,edx
            @@:
                dec eax                       ;Decrement the string
                dec ecx
                jz @F
                cmp byte ptr [eax],'\'
                jne @B
                inc eax
            @@:
            ;
            Invoke lstrcpy,addr TrayBalloonNote.szInfoTitle,eax
            Invoke lstrcpy,addr TrayBalloonNote.szInfo,addr szSolvedJpgPath ;szSolvedJpgComments
            ;Invoke MessageBox,NULL,addr TrayBalloonNote.szInfo,addr TrayBalloonNote.szInfoTitle,NULL
            mov TrayBalloonNote.uTimeout,2000
            mov TrayBalloonNote.dwInfoFlags,NIIF_INFO ;NIIF_NONE NIIF_INFO NIIF_WARNING NIIF_ERROR
            Invoke Shell_NotifyIcon,NIM_MODIFY,addr TrayBalloonNote
            ;************************************************
        .endif
        ;
        ;##########################################################
        ;                     SEND TO OPEN
        ;##########################################################
        szText szShellExecuteOpen,"open"
        szText szShellExecuteExplorer,"explorer.exe" ;"Explorer"
        .if dSendArchiveTo==RUNTIME_CAPTURE_SENT_TO_OPEN
            Invoke ShellExecute,hWinMain,NULL,addr szImageFilePath,NULL,NULL,SW_SHOWDEFAULT
            .if eax > 32 ;If the function succeeds, it returns a value greater than 32
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                szText szActivity_StartScreenShot_SendToOpenSucsess,"Opening target file: Success."
                Invoke UpdateActivityLog,addr szActivity_StartScreenShot_SendToOpenSucsess,NULL,1,0
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            .else
                Invoke GetLastError
                Invoke aaerrortext,eax
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                szText szActivity_StartScreenShot_SendToOpenError,"Error Opening target file: "
                Invoke UpdateActivityLog,addr szActivity_StartScreenShot_SendToOpenError,eax,1,0
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            .endif
        .endif
        ;##########################################################
        ;                     SEND TO MAIL
        ;##########################################################
        .if dSendArchiveTo==RUNTIME_CAPTURE_SENT_TO_MAIL
            ;https://learn.microsoft.com/en-us/windows/win32/shell/launch
            ;With ShellExecute you cannot send the attachment and you are also limited to 255 characters for the mail body.
            ;https://microsoft.public.vc.language.narkive.com/9t8DTkYQ/mailto-with-attachment-problem
;            If you want to send a file containing non-printable characters, you
;            need to encode it with something like MIME, uuencode, etc. E.g.:
;
;            --yourdelimiter
;            Content-Type: application/octet-stream; name="somefile.zip"
;            Content-Transfer-Encoding: base64
;            Content-Disposition: attachment; filename="somefile.zip"
;
;            // MIME-encoded data here
;
;            --yourdelimiter--
            ;
            szText szShellExecuteMailTo,"mail"
            szText szShellExecuteEmailTo,"mailto:somebody@email.com?subject=Subject: Please see the enclosed JPG &body=For reference, I've appended" ;"… &attach="
            Invoke lstrcpy,addr szShellExecuteJpgPath,addr szShellExecuteEmailTo
            Invoke ShellExecute,NULL,NULL,addr szShellExecuteJpgPath,NULL,NULL,SW_SHOWDEFAULT
            ;Invoke ShellExecute,hWinMain,NULL,addr szShellExecuteJpgPath,NULL,NULL,SW_SHOWDEFAULT
            .if eax > 32 ;If the function succeeds, it returns a value greater than 32
                Invoke Sleep,2000
                Invoke SendKey,VK_CONTROL,10,0,TRUE,FALSE
                ;
                Invoke SendKey,VK_V,10,0,TRUE,TRUE ;Invoke SendKey,VK_SNAPSHOT,10,100,TRUE,TRUE
                ;
                Invoke SendKey,VK_CONTROL,10,0,FALSE,TRUE
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                szText szActivity_StartScreenShot_SendToMailSucsess,"Opening email client: Success."
                Invoke UpdateActivityLog,addr szActivity_StartScreenShot_SendToMailSucsess,NULL,1,0
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            .else
                Invoke GetLastError
                Invoke aaerrortext,eax
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                szText szActivity_StartScreenShot_SendToMailError,"Error Opening email client: "
                Invoke UpdateActivityLog,addr szActivity_StartScreenShot_SendToMailError,eax,1,0
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            .endif
            ;Invoke ShellExecute,hWinMain,addr szShellExecuteOpen,addr szShellExecuteEmailTo,NULL,NULL,SW_SHOWDEFAULT
        .endif
        ;##########################################################
        ;                     SEND TO LOCATION
        ;##########################################################
        .if bOpenLastCaptureLocationPath==TRUE || dSendArchiveTo==RUNTIME_CAPTURE_SENT_TO_LOCATION
            ;"explore" works as a verb for folders, and NULL does the default (which is Windows explorer for folders).
            ;The traditional name for this function is OpenFileDefaultViewer(), which you can google (meant more for "open" than "explore").
            ;It should work on any file type too, based on what's mapped in the registry for the file extension (eg. readme.txt --> notepad).
            ;szText szShellExecuteOpen,"Open" ;"Explorer" ;"Explorer.exe"
            ;ShellExecute(0,NULL, 'explorer.exe', '/select,C:\WINDOWS\explorer.exe', nil,SW_SHOWNORMAL)
            ;ShellExecute(Handle, 'OPEN',pchar('explorer.exe'),pchar('/select, "' + FileName + '"'), nil, SW_NORMAL);
            szText szShellExecuteSelect,"/select, "
            Invoke lstrcpy,addr szShellExecuteJpgPath,addr szShellExecuteSelect
            Invoke lstrcat,addr szShellExecuteJpgPath,addr szQuote
            Invoke lstrcat,addr szShellExecuteJpgPath,addr szSolvedJpgPath
            Invoke lstrcat,addr szShellExecuteJpgPath,addr szQuote
            Invoke ShellExecute,NULL,addr szShellExecuteOpen,addr szShellExecuteExplorer,addr szShellExecuteJpgPath,NULL,SW_SHOWDEFAULT
            .if eax > 32 ;If the function succeeds, it returns a value greater than 32
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                szText szActivity_StartScreenShot_SendToLocationSucsess,"Opening target location path: Success."
                Invoke UpdateActivityLog,addr szActivity_StartScreenShot_SendToLocationSucsess,NULL,1,0
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            .else
                Invoke GetLastError
                Invoke aaerrortext,eax
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                szText szActivity_StartScreenShot_SendToLocationError,"Error Opening target location path: "
                Invoke UpdateActivityLog,addr szActivity_StartScreenShot_SendToLocationError,eax,1,0
                ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
            .endif
            ;
;            Invoke lstrcpy,addr szShellExecuteJpgPath,addr szSolvedJpgPath
;            ;********** Get Only the Folder name
;            Invoke lstrlen, addr szSolvedJpgPath
;            lea ecx,szSolvedJpgPath
;            add eax,ecx
;            @@:
;            dec eax                       ;Decrement the string
;            cmp byte ptr [eax],'\'
;            jne @B
;            inc eax
;            mov byte ptr [eax],0
;            Invoke lstrcpy,addr szShellExecuteJpgPath,ecx
;            ;"explore" works as a verb for folders, and NULL does the default (which is Windows explorer for folders).
;            Invoke ShellExecute,NULL,addr szShellExecuteOpen,addr szShellExecuteJpgPath,NULL,NULL,SW_SHOWDEFAULT
            ;
        .endif
        jmp ExitShot
        ;
    ErrorShot:
        Invoke aaerrortext,eax
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        szText szActivity_StartScreenShot_EncoderError,"File Encoder error: "
        Invoke UpdateActivityLog,addr szActivity_StartScreenShot_EncoderError,eax,1,0
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        szText szActivity_StartScreenshot_TitleError,9,9,9,"*** Screen capture call failed ****"
        Invoke UpdateActivityLog,addr szActivity_StartScreenshot_TitleError,NULL,2,0
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    ExitShot:
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ACTIVITY LOG ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        szText szActivity_StartScreenShot_EncoderExits,9,9,9,"*** The Screen capture call ends successfully ***"
        Invoke UpdateActivityLog,addr szActivity_StartScreenShot_EncoderExits,eax,2,0
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    xor eax, eax
    ret
TakeScreenShot                  endp
align 4
StartAutoShotLoop               proc bSendArchiveTo:BYTE,bMinutes:BYTE,bSeconds:BYTE,bOnlyOneTime:BYTE
                                ;
                                ;CALLED from;  Winmain WM_INITDIALOG, "Autoshot" Button or from a "wait 5 seconds to" button
                                ;
                                LOCAL ASL:ASL_PARAM
                                ;ASL_PARAM                   struct
                                ;    bSendArchiveTo              db ?
                                ;    bMinutes                    db ?
                                ;    bSeconds                    db ?
                                ;    bOnlyOneTime                db ?
                                ;ASL_PARAM                   ends
                                ;sAutoShotLoop_Param        ALS_PARAM <?>
            .if bOnlyOneTime == FALSE && hAutoShotLoop_ThreadHandle !=NULL
                ;Terminate previous thread to start new one.
            .endif
            mov ASL.bMinutes,0
            mov al,bSeconds
            mov ASL.bSeconds,al
            mov al,bSendArchiveTo
            mov ASL.bSendArchiveTo,al
            mov al,bOnlyOneTime
            mov ASL.bOnlyOneTime,al
            lea edx,ASL
            mov ecx,DWORD PTR [edx]
            Invoke CreateThread,NULL,NULL,offset AutoShotLoop,ecx,0,addr dAutoShotLoop_ThreadID
            ;If the function succeeds, the return value is a handle to the new thread.
            .if eax==NULL
                jmp Exit_StartAutoShotLoop
            .endif
            mov hAutoShotLoop_ThreadHandle,eax ;If the function succeeds, the return value is a handle to the new thread.
            Invoke SetThreadPriority,eax,THREAD_MODE_BACKGROUND_BEGIN ;THREAD_PRIORITY_ABOVE_NORMAL
            .if bOnlyOneTime == FALSE
                mov dAutoShotLoop_IsRunning,TRUE
                ;Update CHECK buttong if called form WM_INITDIALOG
                ;Invoke GetDlgItem,hWinMain,IDC_BTN_AUTOSHOT
                ;Invoke SendMessage,eax,BM_SETCHECK,BST_CHECKED,NULL
            .endif
            mov eax,hAutoShotLoop_ThreadHandle
        Exit_StartAutoShotLoop:
    ret
StartAutoShotLoop               endp
align 4
AutoShotLoop                    proc dSendArchiveTo_Minutes_Seconds_OnlyOneTime:DWORD
                                LOCAL bLocalMinutes:BYTE
                                LOCAL bLocalSeconds:BYTE
                                LOCAL bLocalSendArchiveTo:BYTE
                                LOCAL bLocalOnlyOneTime:BYTE
                                LOCAL dMilliSeconds:DWORD

                                LOCAL ASL_Param:DWORD
                                ;
                                ;ASL_PARAM                   struct
                                ;    bSendArchiveTo              db ?
                                ;    bMinutes                    db ?
                                ;    bSeconds                    db ?
                                ;    bOnlyOneTime                db ?
                                ;ASL_PARAM                   ends
                                ;sAutoShotLoop_Param        ALM_PARAM <?>
        ;
        mov eax,dSendArchiveTo_Minutes_Seconds_OnlyOneTime
        mov ASL_Param,eax
        lea edx,ASL_Param
        assume edx:PTR ASL_PARAM
            mov al,[edx].bSendArchiveTo
            mov bLocalSendArchiveTo,al
            mov al,[edx].bMinutes
            mov bLocalMinutes,al
            mov al,[edx].bSeconds
            mov bLocalSeconds,al
            mov al,[edx].bOnlyOneTime
            mov bLocalOnlyOneTime,al
        assume edx:nothing
        ;================================================
        ;     Seconds to Milliseconds
        ;================================================
        ; Get Milliseconds to timer
        movzx eax,bLocalMinutes
        mov ecx,60
        xor edx,edx
        mul ecx                     ;eax = Minutes * 60
        movzx ecx,bLocalSeconds
        add eax,ecx                 ;eax = Total Seconds
        mov ecx,1000
        xor edx,edx
        mul ecx                     ;eax = 1000*(Minutes*Seconds)
        ;
        mov dMilliSeconds,eax
        ;
        ;
    AutoShotLoop_NextShot:
        .if bLocalOnlyOneTime==TRUE
            ;************************************
            ;      WAIT 5 SECONDS FOR OnlyOneTime
            ;************************************
            Invoke Sleep,dMilliSeconds
        .endif
        ;
        Invoke TakeScreenShot,bCaptureWindow,bLocalSendArchiveTo
        ;
        .if bLocalOnlyOneTime==TRUE || dMilliSeconds == 0
            Invoke ExitThread,eax
            jmp Exit_AutoShotLoop
        .endif
        ;************************************
        ;   Check if Error Code
        ;************************************
;        .if eax
;            mov ecx,eax
;            jmp AutoShotLoop_ErrorShot
;        .endif
        ;
        ;************************************
        ;      WAIT UNTIL NEXT LAPSE
        ;************************************
        Invoke Sleep,dMilliSeconds
        ;************************************
        ;   Check if Thread still active
        ;************************************
        .if dAutoShotLoop_IsRunning==TRUE
            jmp AutoShotLoop_NextShot
        .else
            jmp Exit_AutoShotLoop
        .endif
        ;
        ;
    AutoShotLoop_ErrorShot:
            ;TODO: Uncheck AutoShot Button
            mov dAutoShotLoop_IsRunning,FALSE
            Invoke ExitThread,eax
            jmp Exit_AutoShotLoop
    Exit_AutoShotLoop:
        xor eax,eax
    Return_AutoShotLoop:
        ;
    ret
AutoShotLoop                    endp
align 4
StopAutoShotLoop                proc
                                ;
        .if dAutoShotLoop_IsRunning==TRUE
            Invoke CloseHandle,hAutoShotLoop_ThreadHandle
            .if eax==NULL ;If the function succeeds, the return value is nonzero.
                ;Invoke TerminateThread, Valid Thread?
            .else
                mov hAutoShotLoop_ThreadHandle,NULL
                mov dAutoShotLoop_IsRunning,FALSE
            .endif
        .endif
    ret
StopAutoShotLoop                endp
align 4
UpdateActivityLog               proc pActivityTitle:DWORD,pActivityText:DWORD,bNewLinesBeforeTitle:BYTE,bNewLinesAfterTitle:BYTE
                                LOCAL dNewLineCount:DWORD;
                                ;
        .if pActivityTitle==NULL && pActivityText==NULL ;Clear Activity Log
            Invoke FillBuffer,addr szActivityLogText,sizeof szActivityLogText,0
            .if hWinMain
                Invoke SetDlgItemText,hWinMain,IDC_EDITBOX_ACTIVITY_LOG,addr szActivityLogText
            .endif
            jmp Exit_UpdateActivityLog
        .endif
        .if pActivityTitle==-1 && pActivityText==-1 ;Before Open Window: Init Activity Log
            Invoke SetDlgItemText,hWinMain,IDC_EDITBOX_ACTIVITY_LOG,addr szActivityLogText
            jmp Exit_UpdateActivityLog
        .endif
        ;
        .if hWinMain
            Invoke GetDlgItemText,hWinMain,IDC_EDITBOX_ACTIVITY_LOG,addr szActivityLogText,sizeof szActivityLogText
        .endif
        ;
        .if bNewLinesBeforeTitle
            movzx eax,bNewLinesBeforeTitle
            mov dNewLineCount,eax
            @@:
                cmp dNewLineCount,0
                je @F
                Invoke lstrcat,addr szActivityLogText,addr szNewLine
                dec dNewLineCount
                jmp @B
            @@:
        .endif
        .if pActivityTitle
            Invoke lstrcat,addr szActivityLogText,pActivityTitle
        .endif
         .if bNewLinesAfterTitle
            movzx eax,bNewLinesBeforeTitle
            mov dNewLineCount,eax
            @@:
                cmp dNewLineCount,0
                je @F
                Invoke lstrcat,addr szActivityLogText,addr szNewLine
                dec dNewLineCount
                jmp @B
            @@:
        .endif
        .if pActivityText
            Invoke lstrcat,addr szActivityLogText,pActivityText
        .endif
        .if hWinMain
            Invoke SetDlgItemText,hWinMain,IDC_EDITBOX_ACTIVITY_LOG,addr szActivityLogText
        .endif
        Exit_UpdateActivityLog:
    ret
UpdateActivityLog               endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
;                       IniManager  PROCEDURE
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
IniManager                      proc  hWnd:DWORD, bAction:BYTE ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
                        ;=====================
                        ; Put LOCALs on stack
                        ;=====================
                        LOCAL  szSingleDigitValue[2]:BYTE
                        LOCAL  bLocalOperation:BYTE
                        LOCAL  dIDC_CHKx:DWORD
                        ;************ Registry ***************
                        LOCAL  hReg:DWORD
                        LOCAL  dSize:DWORD
                        LOCAL  dDWORD:DWORD
                        LOCAL  szTEXTO[MAX_PATH]:BYTE
                        LOCAL  bUsarRegistro:BYTE
                        szText szRegistryCarabeZ,"Software\CarabeZ\CapturA+"
                        szText szREGSZ,"REG_SZ"
                        szText szREGDWORD,"REG_DWORD"
                        ;************ General Setting ***************
                        szText SectApp,"CapturA+"
                        szText KeyCaptureWindow,"CaptureWindow"
                        szText KeyCaptureLeft,"CaptureLeft"
                        szText KeyCaptureTop,"CaptureTop"
                        szText KeyCaptureRight,"CaptureRight"
                        szText KeyCaptureBottom,"CaptureBottom"
                        szText KeyCaptureCursor,"CaptureCursor"
                        szText KeyCaptureMode,"CaptureMode"
                        szText KeyMouseClickSnapDelay,"MouseClickSnapDelay"
                        szText KeyTimerMinutes,"AutoShotTimerMinutes"
                        szText KeyTimerSeconds,"AutoShotTimerSeconds"
                        szText KeyImageMode,"ImageMode"
                        szText KeyImageQuality,"ImageQuality"
                        szText KeyImageZoom,"ImageZoom"
                        szText KeySnapShotCounter,"SnapShotCounter"
                        szText KeyAutoShotStartAtInit,"AutoShotStartAtInit"
                        szText KeyCopyScreenshotToClipboard,"CopyScreenshotToClipboard"
                        ;
                        szText KeyRunWhenWindowsStarts,"RunWhenWindowsStarts"
                        szText KeyHideApplicationWindowWhenOpening,"HideApplicationWindowWhenOpening"
                        szText KeyAudibleNotificationAfterCapture,"AudibleNotificationAfterCapture"
                        szText KeyOpenLastCaptureLocationPath,"OpenLastCaptureLocationPath"
                        szText KeyShowANotificationIcon,"ShowANotificationIcon"
                        szText KeyDisplayANotificationCapture,"DisplayANotificationCapture"
                        szText KeyMinimizeWindowIntoNotificationIcon,"MinimizeWindowIntoNotificationIcon"
                        ;
                        szText KeyGeneralSetting,"BitwiseGeneralSetting"
                        szText KeyTargetType,"TargetType"
                        szText KeyArchiveMode,"ArchiveMode"
                        szText KeyArchivePath,"ArchivePath"
                        ;szText KeyArchivePathDefault,"%TEMPDIR%\Snaps_%DATE%\Cap_%COUNTER%.jpg"
                        szText KeyArchivePathDefault,"%TEMP%\Snaps_%YEAR%%MONTH%\Cap_%COUNTER%.jpg"
                        szText KeyJpgComments,"JpgComment"
                        szText KeyJpgCommentsDefault,"©%YEAR% www.carabez.com"
                        ;
    ;
    Invoke GetRegistryMouseMenu,addr szMouseHoverTime,addr szMenuShowDelay,addr szDoubleClickSpeed
    Invoke a2dw,addr szDoubleClickSpeed
    mov dDoubleClickSpeed,eax
    ;
    ;Invoke RegOpenKeyEx,HKEY_CURRENT_USER,addr szRegistryCarabeZ,0,KEY_ALL_ACCESS,addr hReg ;HKEY_CURRENT_USER HKEY_LOCAL_MACHINE
    Invoke RegCreateKeyEx,HKEY_CURRENT_USER,addr szRegistryCarabeZ,0, NULL,0,KEY_ALL_ACCESS, NULL, addr hReg,NULL
    .if eax==ERROR_SUCCESS
        mov bUsarRegistro,TRUE
        ;Invoke MessageBeep,NULL
    .else
        mov bUsarRegistro,FALSE
        xor eax,eax
        mov hReg,eax
    .endif
    .if bUsarRegistro==TRUE
        .if bIniArchiveExists==TRUE
            mov bUsarRegistro,FALSE
        .endif
    .endif
    ;
    .if bAction == 0 ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: TargetType ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyTargetType,2,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyTargetType,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        cmp eax,0
        jg @F
            mov al,1
        @@:
        cmp eax,2
        jle @F
            mov al,1
        @@:
        mov bTargetType,al
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: CaptureWindow ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyCaptureWindow,1,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyCaptureWindow,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
                inc eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        cmp eax,0
        jg @F
            mov al,1
        @@:
        cmp eax,3
        jle @F
            mov al,1
        @@:
        mov bCaptureWindow,al
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyCaptureLeft,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyCaptureLeft,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov rcCapture.left,eax
        ;
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyCaptureTop,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyCaptureTop,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov rcCapture.top,eax
        ;
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyCaptureRight,32,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyCaptureRight,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,32
            .endif
            mov eax,dDWORD
        .endif
        mov rcCapture.right,eax
        ;
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyCaptureBottom,32,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyCaptureBottom,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,32
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov rcCapture.bottom,eax
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: CaptureCursor ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyCaptureCursor,1,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyCaptureCursor,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
                inc eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        cmp eax,0
        jg @F
            mov al,1
        @@:
        cmp eax,3
        jle @F
            mov al,1
        @@:
        mov bCaptureCursor,al
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: ImageMode ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyImageMode,1,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyImageMode,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
                inc eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        cmp eax,0
        jg @F
            mov al,1
        @@:
        cmp eax,2
        jle @F
            mov al,1
        @@:
        mov bImageMode,al
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: CaptureMode ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyCaptureMode,2,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyCaptureMode,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
                inc eax
                inc eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        cmp eax,0
        jg @F
            mov al,1
        @@:
        cmp eax,3
        jle @F
            mov al,1
        @@:
        mov bCaptureMode,al
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: Mouse Click Delay ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyMouseClickSnapDelay,400,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyMouseClickSnapDelay,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS ;Does not Exist?
                Invoke a2dw,addr szMenuShowDelay
                add eax,200
            .else
                mov eax,dDWORD
            .endif
        .endif
        cmp eax,0
        jge @F
            xor eax,eax  ;Negative value? 0 = NO DELAY
        @@:
        cmp eax,MAX_CLICK_DELAY
        jle @F
            mov eax,MAX_CLICK_DELAY
        @@:
        mov dMouseClickSnapDelay,eax
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: TimerMinutes ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyTimerMinutes,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyTimerMinutes,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        cmp eax,0
        jge @F
            xor eax,eax
        @@:
        cmp eax,99
        jle @F
            xor eax,eax
        @@:
        mov bAutoShot_TimerMinutes,al
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: TimerSeconds ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyTimerSeconds,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyTimerSeconds,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
                inc eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        cmp eax,0
        jge @F
            xor eax,eax
        @@:
        cmp eax,59
        jle @F
            xor eax,eax
        @@:
        mov bAutoShot_TimerSeconds,al
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: AutoShot Start At Init ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyAutoShotStartAtInit,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyAutoShotStartAtInit,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,0
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov bAutoShotStartAtInit,al
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: Copy Screenshot To Clipboard ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyCopyScreenshotToClipboard,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyCopyScreenshotToClipboard,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,1
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov bCopyFilePathToClipboard ,al
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: Run When Windows Starts ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyRunWhenWindowsStarts,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyRunWhenWindowsStarts,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,1
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov bRunWhenWindowsStarts,al
        Invoke UpdateRegistryRun,bRunWhenWindowsStarts
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: Hide Application Window When Opening ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyHideApplicationWindowWhenOpening,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyHideApplicationWindowWhenOpening,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,0
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov bHideApplicationWindowWhenOpening,al
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: Audible Notification After Capture ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyAudibleNotificationAfterCapture,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyAudibleNotificationAfterCapture,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,1
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov bAudibleNotificationAfterCapture,al
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: Open Last Capture Location Path ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyOpenLastCaptureLocationPath,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyOpenLastCaptureLocationPath,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,0
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov bOpenLastCaptureLocationPath,al
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: Show A Notification Icon  ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyShowANotificationIcon ,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyShowANotificationIcon ,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,0
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov bShowANotificationIcon,al
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: Display A Notification Capture  ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyDisplayANotificationCapture,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyDisplayANotificationCapture,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,0
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov bDisplayANotificationCapture,al
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: Minimize Window Into Notification Icon  ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyMinimizeWindowIntoNotificationIcon,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyMinimizeWindowIntoNotificationIcon,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,1
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov bCopyScreenShotToClipboard,al
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: ArchiveMode ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyArchiveMode,1,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyArchiveMode,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
                inc eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        cmp eax,0
        jg @F
            mov al,1
        @@:
        cmp eax,2
        jle @F
            mov al,1
        @@:
        mov bArchiveMode,al
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: ArchivePath ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileString,addr SectApp,addr KeyArchivePath,addr KeyArchivePathDefault,addr szArchivePath,MAX_PATH,addr szIniFile
        .else
            Invoke FillBuffer,addr szTEXTO,sizeof szTEXTO,0
            mov dSize,sizeof szTEXTO
            Invoke RegQueryValueEx,hReg,addr KeyArchivePath,0,NULL,addr szTEXTO,addr dSize
            .if eax!=ERROR_SUCCESS
                lea eax,KeyArchivePathDefault
            .else
                lea eax,szTEXTO
            .endif
            Invoke lstrcpy,addr szArchivePath,eax
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: JpgComments ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileString,addr SectApp,addr KeyJpgComments,addr KeyJpgCommentsDefault,addr szJpgComments,MAX_PATH,addr szIniFile
        .else
            Invoke FillBuffer,addr szTEXTO,sizeof szTEXTO,0
            mov dSize,sizeof szTEXTO
            Invoke RegQueryValueEx,hReg,addr KeyJpgComments,0,NULL,addr szTEXTO,addr dSize
            .if eax!=ERROR_SUCCESS
                lea eax,KeyJpgCommentsDefault
            .else
                lea eax,szTEXTO
            .endif
            Invoke lstrcpy,addr szJpgComments,eax
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: ImageQuality ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyImageQuality,80,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyImageQuality,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
                mov eax,80
            .else
                mov eax,dDWORD
            .endif
        .endif
        cmp eax,0
        jg @F
            mov al,80
        @@:
        cmp eax,100
        jle @F
            mov al,80
        @@:
        mov bImageQuality,al
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: ImageZoom ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyImageZoom,100,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyImageZoom,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,100
            .else
                mov eax,dDWORD
            .endif
        .endif
        cmp eax,0
        jg @F
            mov al,100
        @@:
        cmp eax,250
        jle @F
            mov al,100
        @@:
        mov bImageZoom,al
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ READ: SnapShotCounter ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeySnapShotCounter,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeySnapShotCounter,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov dSnapShotCounter,eax
        ;%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        ;%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    .elseif bAction == 1 || bAction == 2;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
        ;%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        ;%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: Capture Window Target¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        ;bCaptureWindow              db ?  ;TARGET TYPE: RUNTIME_TARGET_TYPE_MONITOR,RUNTIME_TARGET_TYPE_ACTIVE_WINDOWS,RUNTIME_TARGET_TYPE_DESKTOP,RUNTIME_TARGET_TYPE_REGION
        Invoke SendMessage, hCombo3, CB_GETCURSEL, 0, 0 ;The return value is the zero-based index of the currently selected item. If no item is selected, it is CB_ERR.
        .if (eax==-1)
            xor eax,eax
        .endif
        inc eax
        mov bCaptureWindow,al
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyCaptureWindow,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyCaptureWindow,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: CaptureCursor ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        ;bCaptureCursor              db ?  ;CAPTURE CURSOR: RUNTIME_CAPTURE_CURSOR_DISABLE,RUNTIME_CAPTURE_CURSOR_SYSTEM_ARROW,RUNTIME_CAPTURE_CURSOR_RED_ARROW
        Invoke SendMessage, hCombo4, CB_GETCURSEL, 0, 0 ;The return value is the zero-based index of the currently selected item. If no item is selected, it is CB_ERR.
        .if (eax==-1)
            xor eax,eax
        .endif
        inc eax
        mov bCaptureCursor,al
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyCaptureCursor,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyCaptureCursor,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: ImageMode ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        ;bImageMode                  db ?  ;IMAGE MODE: RUNTIME_IMAGE_MODE_COLOR, RUNTIME_IMAGE_MODE_GRAY_SCALE
        Invoke SendMessage, hCombo6, CB_GETCURSEL, 0, 0 ;The return value is the zero-based index of the currently selected item. If no item is selected, it is CB_ERR.
        .if (eax==-1)
            xor eax,eax
        .endif
        inc eax
        mov bImageMode,al
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyImageMode,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyImageMode,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: CaptureMode ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        ;bCaptureMode                db ?  ;CAPTURE MODE: RUNTIME_CAPTURE_MODE_AUTOSHOT,RUNTIME_CAPTURE_MODE_MANUAL,RUNTIME_CAPTURE_MODE_CLICKS
        Invoke SendMessage, hCombo5, CB_GETCURSEL, 0, 0 ;The return value is the zero-based index of the currently selected item. If no item is selected, it is CB_ERR.
        .if (eax==-1)
            xor eax,eax
        .endif
        inc eax
        mov bCaptureMode,al
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyCaptureMode,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyCaptureMode,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: AutoShot Start At Init ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItem,hWnd,IDC_CHK_START_AUTOSHOT_AT_INIT
        Invoke  SendMessage, eax, BM_GETCHECK, 0, 0
        .if eax==BST_CHECKED
            mov bAutoShotStartAtInit,TRUE
        .else
            mov bAutoShotStartAtInit,FALSE
        .endif
        movzx eax,bAutoShotStartAtInit
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyAutoShotStartAtInit,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyAutoShotStartAtInit,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: Copy Screenshot To Clipboard ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItem,hWnd,IDC_CHK_COPY_PATH_TO_CLIPBOARD
        Invoke  SendMessage, eax, BM_GETCHECK, 0, 0
        .if eax==BST_CHECKED
            mov bCopyFilePathToClipboard,TRUE
        .else
            mov bCopyFilePathToClipboard,FALSE
        .endif
        movzx eax,bCopyFilePathToClipboard
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyCopyScreenshotToClipboard,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyCopyScreenshotToClipboard,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: Run When Windows Starts ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItem,hWnd,IDC_CHK_WIN_STARTUP_RUN
        Invoke  SendMessage, eax, BM_GETCHECK, 0, 0
        .if eax==BST_CHECKED
            mov bRunWhenWindowsStarts,TRUE
        .else
            mov bRunWhenWindowsStarts,FALSE
        .endif
        movzx eax,bRunWhenWindowsStarts
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyRunWhenWindowsStarts,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyRunWhenWindowsStarts,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        movzx eax,bRunWhenWindowsStarts
        Invoke UpdateRegistryRun,eax
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE:  Hide Application Window When Opening ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItem,hWnd,IDC_CHK_HIDE_WIN_WHEN_OPENING
        Invoke  SendMessage, eax, BM_GETCHECK, 0, 0
        .if eax==BST_CHECKED
            mov bHideApplicationWindowWhenOpening,TRUE
        .else
            mov bHideApplicationWindowWhenOpening,FALSE
        .endif
        movzx eax,bHideApplicationWindowWhenOpening
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyHideApplicationWindowWhenOpening,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyHideApplicationWindowWhenOpening,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE:  Audible Notification After Capture ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItem,hWnd,IDC_CHK_PLAY_AUDIBLE_SHOT
        Invoke  SendMessage, eax, BM_GETCHECK, 0, 0
        .if eax==BST_CHECKED
            mov bAudibleNotificationAfterCapture,TRUE
        .else
            mov bAudibleNotificationAfterCapture,FALSE
        .endif
        movzx eax,bAudibleNotificationAfterCapture
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyAudibleNotificationAfterCapture,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyAudibleNotificationAfterCapture,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE:  Open Last Capture Location Path ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItem,hWnd,IDC_CHK_OPEN_LAST_CAPTURE_PATH
        Invoke  SendMessage, eax, BM_GETCHECK, 0, 0
        .if eax==BST_CHECKED
            mov bOpenLastCaptureLocationPath,TRUE
        .else
            mov bOpenLastCaptureLocationPath,FALSE
        .endif
        movzx eax,bOpenLastCaptureLocationPath
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyOpenLastCaptureLocationPath,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyOpenLastCaptureLocationPath,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE:  Show A Notification Icon ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItem,hWnd,IDC_CHK_SHOW_NOTIFICATION_ICON
        Invoke  SendMessage, eax, BM_GETCHECK, 0, 0
        .if eax==BST_CHECKED
            mov bShowANotificationIcon ,TRUE
        .else
            mov bShowANotificationIcon,FALSE
        .endif
        movzx eax,bShowANotificationIcon
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyShowANotificationIcon,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyShowANotificationIcon,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE:  Display A Notification Capture ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItem,hWnd,IDC_CHK_DISPLAY_NOTIFICATION_CAPTURE
        Invoke  SendMessage, eax, BM_GETCHECK, 0, 0
        .if eax==BST_CHECKED
            mov bDisplayANotificationCapture,TRUE
        .else
            mov bDisplayANotificationCapture,FALSE
        .endif
        movzx eax,bDisplayANotificationCapture
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyDisplayANotificationCapture,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyDisplayANotificationCapture,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE:  Copy Screenshot Image to Clipboard ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItem,hWnd,IDC_CHK_COPY_SCREENSHOT_CLIPBOARD
        Invoke  SendMessage, eax, BM_GETCHECK, 0, 0
        .if eax==BST_CHECKED
            mov bCopyScreenShotToClipboard,TRUE
        .else
            mov bCopyScreenShotToClipboard,FALSE
        .endif
        movzx eax,bCopyScreenShotToClipboard
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyMinimizeWindowIntoNotificationIcon,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyMinimizeWindowIntoNotificationIcon,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: TargetType ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke SendMessage, hCombo1, CB_GETCURSEL, 0, 0 ;The return value is the zero-based index of the currently selected item. If no item is selected, it is CB_ERR.
        .if (eax==-1)
            xor eax,eax
        .endif
        inc eax
        mov bTargetType,al
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyTargetType,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyTargetType,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: ArchiveMode ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke SendMessage, hCombo2, CB_GETCURSEL, 0, 0 ;The return value is the zero-based index of the currently selected item. If no item is selected, it is CB_ERR.
        .if (eax==-1)
            xor eax,eax
        .endif
        inc eax
        mov bArchiveMode,al
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyArchiveMode,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyArchiveMode,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: ArchivePath ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItemText,hWnd,IDC_EDITBOX_ARCHIVE_PATH,addr szArchivePath,sizeof szArchivePath
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke ValidateIniValue,addr szArchivePath,sizeof szArchivePath
                Invoke WritePrivateProfileString,addr SectApp,addr KeyArchivePath,addr szArchivePath,addr szIniFile
            .else
                Invoke RegSetValueEx,hReg,addr KeyArchivePath,0,REG_SZ,addr szArchivePath,eax
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: JpgComments ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItemText,hWnd,IDC_EDITBOX_JPG_COMMENTS,addr szJpgComments,sizeof szJpgComments
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke ValidateIniValue,addr szArchivePath,sizeof szArchivePath
                Invoke WritePrivateProfileString,addr SectApp,addr KeyJpgComments,addr szJpgComments,addr szIniFile
            .else
                Invoke RegSetValueEx,hReg,addr KeyJpgComments,0,REG_SZ,addr szJpgComments,eax
            .endif
        .endif
        ;
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: Mouse Click Delay ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItemInt,hWnd,IDC_EDITBOX_CLICK_DELAY,NULL,FALSE
        mov dMouseClickSnapDelay,eax
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyMouseClickSnapDelay,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyMouseClickSnapDelay,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: TimerMinutes ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItemInt,hWnd,IDC_EDITBOX_TIMELAPSE_MIN,NULL,FALSE
        mov bAutoShot_TimerMinutes,al
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyTimerMinutes,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyTimerMinutes,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: TimerSeconds ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItemInt,hWnd,IDC_EDITBOX_TIMELAPSE_SEC,NULL,FALSE
        mov bAutoShot_TimerSeconds,al
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyTimerSeconds,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyTimerSeconds,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: ImageQuality ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItemInt,hWnd,IDC_EDITBOX_IMAGE_QUALITY,NULL,FALSE
        mov bImageQuality,al
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyImageQuality,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyImageQuality,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: ImageZoom ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItemInt,hWnd,IDC_EDITBOX_IMAGE_ZOOM,NULL,FALSE
        mov bImageZoom,al
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,eax
                Invoke WritePrivateProfileString,addr SectApp,addr KeyImageZoom,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeyImageZoom,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WRITE: SnapShotCounter ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        Invoke GetDlgItemInt,hWnd,IDC_EDITBOX_COUNTER,NULL,FALSE
        mov dSnapShotCounter,eax
        .if bAction == 1
            .if bUsarRegistro==FALSE
                Invoke wsprintf,addr szBuffer,addr szIntFormat,dSnapShotCounter
                Invoke WritePrivateProfileString,addr SectApp,addr KeySnapShotCounter,addr szBuffer,addr szIniFile
            .else
                mov dDWORD,eax
                Invoke RegSetValueEx,hReg,addr KeySnapShotCounter,0,REG_DWORD,addr dDWORD,sizeof dDWORD
            .endif
        .endif
        ;
    .elseif bAction == 4 ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyCaptureLeft,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyCaptureLeft,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov rcCapture.left,eax
        ;
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyCaptureTop,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyCaptureTop,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov rcCapture.top,eax
        ;
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyCaptureRight,32,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyCaptureRight,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,32
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov rcCapture.right,eax
        ;
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeyCaptureBottom,32,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeyCaptureBottom,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                mov eax,32
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov rcCapture.bottom,eax
    .elseif bAction == 5 ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
        .if bUsarRegistro==FALSE
            Invoke wsprintf,addr szBuffer,addr szIntFormat,rcCapture.left
            Invoke WritePrivateProfileString,addr SectApp,addr KeyCaptureLeft,addr szBuffer,addr szIniFile
        .else
            Invoke RegSetValueEx,hReg,addr KeyCaptureLeft,0,REG_DWORD,addr rcCapture.left,sizeof rcCapture.left
        .endif
        ;
        .if bUsarRegistro==FALSE
            Invoke wsprintf,addr szBuffer,addr szIntFormat,rcCapture.top
            Invoke WritePrivateProfileString,addr SectApp,addr KeyCaptureTop,addr szBuffer,addr szIniFile
        .else
            Invoke RegSetValueEx,hReg,addr KeyCaptureTop,0,REG_DWORD,addr rcCapture.top,sizeof rcCapture.top
        .endif
        ;
        .if bUsarRegistro==FALSE
            Invoke wsprintf,addr szBuffer,addr szIntFormat,rcCapture.right
            Invoke WritePrivateProfileString,addr SectApp,addr KeyCaptureRight,addr szBuffer,addr szIniFile
        .else
            Invoke RegSetValueEx,hReg,addr KeyCaptureRight,0,REG_DWORD,addr rcCapture.right,sizeof rcCapture.right
        .endif
        ;
        .if bUsarRegistro==FALSE
            Invoke wsprintf,addr szBuffer,addr szIntFormat,rcCapture.bottom
            Invoke WritePrivateProfileString,addr SectApp,addr KeyCaptureBottom,addr szBuffer,addr szIniFile
        .else
            Invoke RegSetValueEx,hReg,addr KeyCaptureBottom,0,REG_DWORD,addr rcCapture.bottom,sizeof rcCapture.bottom
        .endif

    .elseif  bAction == 6 ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
        .if bUsarRegistro==FALSE
            Invoke GetPrivateProfileInt, addr SectApp,addr KeySnapShotCounter,0,addr szIniFile
        .else
            mov dSize,sizeof dDWORD
            Invoke RegQueryValueEx,hReg,addr KeySnapShotCounter,0,NULL,addr dDWORD,addr dSize
            .if eax!=ERROR_SUCCESS
                xor eax,eax
            .else
                mov eax,dDWORD
            .endif
        .endif
        mov dSnapShotCounter,eax
    .elseif  bAction == 7 ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
        Invoke GetDlgItemInt,hWnd,IDC_EDITBOX_COUNTER,NULL,FALSE
        mov dSnapShotCounter,eax
        .if bUsarRegistro==FALSE
            Invoke wsprintf,addr szBuffer,addr szIntFormat,dSnapShotCounter
            Invoke WritePrivateProfileString,addr SectApp,addr KeySnapShotCounter,addr szBuffer,addr szIniFile
        .else
            mov dDWORD,eax
            Invoke RegSetValueEx,hReg,addr KeySnapShotCounter,0,REG_DWORD,addr dDWORD,sizeof dDWORD
        .endif
    .endif ;.if bAction==0 ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
    .if hReg!=NULL
        Invoke RegCloseKey,hReg
    .endif
    ret
IniManager                      endp
align 4
ValidateIniValue                proc pSource:DWORD,dSize:DWORD
                                ;
                                LOCAL       szValidBuffer[4096]:BYTE
                                LOCAL       bCopyBack:DWORD
                                LOCAL       bSameSize:DWORD
                                LOCAL       bAddExtra:DWORD
                                szText      szQuoteSimple,"'"
                                szText      szQuoteDoble,'"'
                                ;Converts "C:\MyApp\" to ""C:\MyApp\"" or
                                ;Converts 'My quote text' to ''My quote text''
                                ;so when stored in ini file can be retrieve back to single quotes.
                                ;https://en.wikipedia.org/wiki/INI_file
                                ;
        Invoke FillBuffer,addr szValidBuffer,sizeof szValidBuffer,0
        mov bSameSize,FALSE
        mov bCopyBack,FALSE
        mov bAddExtra,FALSE
        invoke lstrlen,pSource
        .if eax==0
            jmp ExitValidateIniValue
        .elseif eax > 4095 ;Larger than buffer?
            mov eax,-1
            jmp ExitValidateIniValue
        .endif

        invoke lstrlen,pSource
        mov ecx,eax
        mov edx,pSource
        .if byte ptr [edx] == '"'
            .if byte ptr [edx+ecx-1] == '"'
                invoke lstrcpy,addr szValidBuffer,addr szQuoteDoble
                invoke lstrcat,addr szValidBuffer,pSource
                invoke lstrcat,addr szValidBuffer,addr szQuoteDoble
                invoke lstrcpy,pSource,addr szValidBuffer
            .else
                ;invoke lstrcpy,addr szValidBuffer,addr szQuoteDoble
                ;invoke lstrcat,addr szValidBuffer,pSource
                ;invoke lstrcpy,pSource,addr szValidBuffer
            .endif
        .elseif byte ptr [edx+ecx-1] == '"'
            ;invoke lstrcpy,addr szValidBuffer,pSource
            ;invoke lstrcat,addr szValidBuffer,addr szQuoteDoble
            ;invoke lstrcpy,pSource,addr szValidBuffer
        .endif
        invoke lstrlen,pSource
        mov ecx,eax
        mov edx,pSource
        .if byte ptr [edx] == "'"
            .if byte ptr [edx+ecx-1] == "'"
                invoke lstrcpy,addr szValidBuffer,addr szQuoteSimple
                invoke lstrcat,addr szValidBuffer,pSource
                invoke lstrcat,addr szValidBuffer,addr szQuoteSimple
                invoke lstrcpy,pSource,addr szValidBuffer
            .else
                ;invoke lstrcpy,addr szValidBuffer,addr szQuoteDoble
                ;invoke lstrcat,addr szValidBuffer,pSource
                ;invoke lstrcpy,pSource,addr szValidBuffer
            .endif
        .elseif byte ptr [edx+ecx-1] == '"'
            ;invoke lstrcpy,addr szValidBuffer,pSource
            ;invoke lstrcat,addr szValidBuffer,addr szQuoteDoble
            ;invoke lstrcpy,pSource,addr szValidBuffer
        .endif
        jmp ExitValidateIniValue
        ExitValidateIniValue:
        ;Invoke OutputDebugString,pSource
        ;Return 0 on success or Required string size
    ret
ValidateIniValue                endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
;                       Region Dialog
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
RegionDlg                       proc hParent:DWORD,lpIcon:DWORD

    Dialog "Move and resize this window to cover the area you want to capture.","MS Sans Serif",8, \            ; caption,font,pointsize
            WS_OVERLAPPED or WS_SIZEBOX or WS_SYSMENU or DS_CENTER, \     ; style
            1, \                                            ; control count
            0,0,324,128, \                                  ; x y co-ordinates
            1024                                            ; memory buffer size
    DlgButton   "Save",WS_TABSTOP OR 1  OR 8000H,4,4,40,16,IDC_BTN6


    CallModalDialog hInstance,hParent,RegionDlgProc,ADDR hParent

    ret

RegionDlg                       endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
;                       RegionDlg Procedure
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
RegionDlgProc                   proc hWin:DWORD,uMsg:DWORD,wParam:DWORD,lParam:DWORD

      ;
      .if uMsg == WM_INITDIALOG
      ; -----------------------------------------
      ; write the parameters passed in "lParam"
      ; to the dialog's GWL_USERDATA address.
      ; -----------------------------------------
      Invoke SetWindowLong,hWin,GWL_USERDATA,lParam
      mov eax, lParam
      mov eax, [eax+4]
      Invoke SendMessage,hWin,WM_SETICON,1,eax
      ;
      .elseif uMsg == WM_COMMAND
          .if wParam == IDC_BTN6
                Invoke GetWindowRect,hWin,addr rcCapture
                Invoke IniManager,hWin,5   ;(WinHandle),(0=GET,1=PUTMAIN,2=READMAIN,3=PUTSECURITY,4=READREGION,5=PUTREGION,6=READCOUNTER,7=PUTCOUNTER)
                jmp Exit_About
          .endif
      .elseif uMsg == WM_CLOSE
        Exit_About:
        Invoke EndDialog,hWin,0
      .endif

      xor eax, eax
      ret

RegionDlgProc                   endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
;##########################################################################
;                 *** GET Mouse and Menu timing ***
;##########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
GetRegistryMouseMenu            proc pszMouseHoverTime:DWORD,pszMenuShowDelay:DWORD,pszDoubleClickSpeed:DWORD
                        ;
                        LOCAL dResult:DWORD
                        LOCAL hReg:DWORD
                        LOCAL dSize:DWORD
                        LOCAL szValidBuffer[MAX_PATH]:BYTE
                        szText szMouse,"Control Panel\Mouse"
                        szText KeyMouseHoverTime,"MouseHoverTime"
                        szText KeyDoubleClickSpeed,"DoubleClickSpeed"
                        szText szDesktop,"Control Panel\Desktop"
                        szText KeyMenuShowDelay,"MenuShowDelay"
        ;******************************************
        ;Invoke OutputDebugString,addr KeyMouseHoverTime
        ;ret
        xor eax,eax
        mov dResult,eax
        Invoke RegOpenKeyEx,HKEY_CURRENT_USER,addr szMouse,0,KEY_READ,addr hReg
        .if eax==ERROR_SUCCESS
            Invoke FillBuffer,addr szValidBuffer,sizeof szValidBuffer,0
            mov dSize,sizeof szValidBuffer
            Invoke RegQueryValueEx,hReg,addr KeyMouseHoverTime,0,NULL,addr szValidBuffer,addr dSize
            .if eax==ERROR_SUCCESS
                ;Invoke OutputDebugString,addr KeyMouseHoverTime
                ;Invoke OutputDebugString,addr szValidBuffer
                Invoke lstrcpy,pszMouseHoverTime,addr szValidBuffer
                mov dResult,ERROR_SUCCESS
            .endif
            Invoke RegCloseKey,hReg
        .endif
        ;******************************************
        xor eax,eax
        Invoke RegOpenKeyEx,HKEY_CURRENT_USER,addr szDesktop,0,KEY_READ,addr hReg
        .if eax==ERROR_SUCCESS
            Invoke FillBuffer,addr szValidBuffer,sizeof szValidBuffer,0
            mov dSize,sizeof szValidBuffer
            Invoke RegQueryValueEx,hReg,addr KeyMenuShowDelay,0,NULL,addr szValidBuffer,addr dSize
            .if eax==ERROR_SUCCESS
                ;Invoke OutputDebugString,addr KeyMenuShowDelay
                ;Invoke OutputDebugString,addr szValidBuffer
                Invoke lstrcpy,pszMenuShowDelay,addr szValidBuffer
                mov dResult,ERROR_SUCCESS
            .endif
            Invoke RegCloseKey,hReg
        .endif
        ;
        xor eax,eax
        mov dResult,eax
        Invoke RegOpenKeyEx,HKEY_CURRENT_USER,addr szMouse,0,KEY_READ,addr hReg
        .if eax==ERROR_SUCCESS
            Invoke FillBuffer,addr szValidBuffer,sizeof szValidBuffer,0
            mov dSize,sizeof szValidBuffer
            Invoke RegQueryValueEx,hReg,addr KeyDoubleClickSpeed,0,NULL,addr szValidBuffer,addr dSize
            .if eax==ERROR_SUCCESS
                ;Invoke OutputDebugString,addr KeyMouseHoverTime
                ;Invoke OutputDebugString,addr szValidBuffer
                Invoke lstrcpy,pszDoubleClickSpeed,addr szValidBuffer
                mov dResult,ERROR_SUCCESS
            .endif
            Invoke RegCloseKey,hReg
        .endif
        ;******************************************
        .if dResult==ERROR_SUCCESS
            mov eax,ERROR_SUCCESS
        .endif
    ret
GetRegistryMouseMenu            endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
;##########################################################################
;                 *** WRITE KEY IN REGISTER ***
;##########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
UpdateRegistryRun               proc dRunWhenWindowsStart:DWORD
                        ;
                        ;LOCAL     szLocalTempBuffer[MAX_PATH+MAX_PATH]:BYTE
                        ;TODO: ONLY Update if value changed.
                        LOCAL hReg:DWORD
                        szText szRegistryRun,"Software\Microsoft\Windows\CurrentVersion\Run"
    ;******************************************
        Invoke RegOpenKeyEx,HKEY_LOCAL_MACHINE,addr szRegistryRun,0,KEY_ALL_ACCESS,addr hReg
        .if eax==ERROR_SUCCESS
            .if dRunWhenWindowsStart
                Invoke lstrlen,addr szQuotedPath
                Invoke RegSetValueEx,hReg,addr szProgramName,0,REG_SZ,addr szQuotedPath,eax
            .else
                Invoke RegDeleteValue,hReg,addr szProgramName
            .endif
            Invoke RegCloseKey,hReg
        .endif
    ret
UpdateRegistryRun               endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
;##########################################################################
;                 *** WRITE KEY IN REGISTER ***
;##########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
UpdateRegistryPrintScreen       proc dNewValue:DWORD ;0 = Off, 1 = On
                        ;
                        LOCAL dOldValue:DWORD
                        LOCAL hReg:DWORD
                        LOCAL dSize:DWORD
                        LOCAL dDWORD:DWORD
                        szText szKeyboard,"Control Panel\Keyboard"
                        szText KeyPrintScreenKey,"PrintScreenKeyForSnippingEnabled"
        ;******************************************
        xor eax,eax
        mov dOldValue,eax
        Invoke RegOpenKeyEx,HKEY_CURRENT_USER,addr szKeyboard,0,KEY_ALL_ACCESS,addr hReg
        .if eax==ERROR_SUCCESS
            ;Invoke OutputDebugString,addr szKeyboard
            mov dSize,sizeof dOldValue
            Invoke RegQueryValueEx,hReg,addr KeyPrintScreenKey,0,NULL,addr dOldValue,addr dSize
            mov eax,dOldValue ;0 = Off, 1 = On
            .if eax!=dNewValue
                ;Invoke OutputDebugString,addr KeyPrintScreenKey
                Invoke RegSetValueEx,hReg,addr KeyPrintScreenKey,0,REG_DWORD,addr dNewValue,sizeof dNewValue
            .endif
            Invoke RegCloseKey,hReg
        .endif
        mov eax,dOldValue
    ret
UpdateRegistryPrintScreen       endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
;           *** ReplaceMacros **
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
SolveMacros                     proc lpSourceStr:DWORD,lpTargetStr:DWORD

                          LOCAL     szLocalTempBuffer[MAX_PATH]:BYTE
                          LOCAL     szTargetBuffer[MAX_PATH]:BYTE
                          LOCAL     dSize :DWORD
                          LOCAL     hReg:DWORD
                          LOCAL     LocalTime:SYSTEMTIME
                          szText    szTODO,"TODO"
                          szText    szCurrentVersion,"Software\Microsoft\Windows\CurrentVersion"
                          szText    szProgramFilesDir,"ProgramFilesDir"
                          szText    szTwoDigitFormat,"%02d"                 ; To format '1' into '01:'
    ;
    Invoke lstrcpy,addr szTargetBuffer,lpSourceStr ;Source to Temp
    Invoke GetLocalTime, addr LocalTime
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ YEAR ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    xor eax,eax
    mov ax,LocalTime.wYear
    mov dSize,eax
    Invoke wsprintf,addr szLocalTempBuffer,addr szIntFormat,eax
    ;Invoke dwtoa,dSize,addr szLocalTempBuffer
    Invoke ReplaceMacros,addr szTargetBuffer,addr szYEAR,addr szLocalTempBuffer          ;Current year
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ MONTH ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    xor eax,eax
    mov ax,LocalTime.wMonth
    lea ecx,szLocalTempBuffer
    add ecx,4
    Invoke wsprintfA, ecx, addr szTwoDigitFormat,eax
    lea ecx,szLocalTempBuffer
    add ecx,4
    Invoke ReplaceMacros,addr szTargetBuffer,addr szMONTH,ecx                                   ;Month of the Year (1-12)
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ DAY ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    xor eax,eax
    mov ax,LocalTime.wDay
    lea ecx,szLocalTempBuffer
    add ecx,6
    Invoke wsprintfA, ecx, addr szTwoDigitFormat,eax
    lea ecx,szLocalTempBuffer
    add ecx,6
    Invoke ReplaceMacros,addr szTargetBuffer,addr szDAY,ecx                                     ;Day of the month (1-31)
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ DATE ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke ReplaceMacros,addr szTargetBuffer,addr szDATE,addr szLocalTempBuffer          ;Current date (in the format YYYYMMDD)
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ HOUR ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    xor eax,eax
    mov ax,LocalTime.wHour
    Invoke wsprintfA, addr szLocalTempBuffer, addr szTwoDigitFormat,eax
    Invoke ReplaceMacros,addr szTargetBuffer,addr szHOUR,addr szLocalTempBuffer          ;LocalTime Hour 0-24
    ;Invoke MessageBox,NULL,addr szHOUR,addr szLocalTempBuffer,NULL
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ MINUTE ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    xor eax,eax
    mov ax,LocalTime.wMinute
    lea ecx,szLocalTempBuffer
    add ecx,2
    Invoke wsprintfA, ecx, addr szTwoDigitFormat,eax
    lea ecx,szLocalTempBuffer
    add ecx,2
    Invoke ReplaceMacros,addr szTargetBuffer,addr szMINUTE,ecx                                      ;LocalTime Hour 0-59
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ SECOND ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    xor eax,eax
    mov ax,LocalTime.wSecond
    lea ecx,szLocalTempBuffer
    add ecx,4
    Invoke wsprintfA, ecx, addr szTwoDigitFormat,eax
    lea ecx,szLocalTempBuffer
    add ecx,4
    Invoke ReplaceMacros,addr szTargetBuffer,addr szSECOND,ecx                                      ;LocalTime Second 0-59
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ TIME ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke ReplaceMacros,addr szTargetBuffer,addr szTIME,addr szLocalTempBuffer        ;Current time (in the format HHMMSS)
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WEEKDAY ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    xor eax,eax
    mov ax,word ptr [LocalTime.wDayOfWeek]
    inc ax
    mov dSize,eax
    Invoke wsprintf,addr szLocalTempBuffer,addr szIntFormat,eax
    ;Invoke dwtoa,dSize,addr szLocalTempBuffer
    Invoke ReplaceMacros,addr szTargetBuffer,addr szWEEKDAY,addr szLocalTempBuffer     ;Days since Sunday (1 – 7)
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ DRIVE ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    lea eax,szProgramPath
    mov cl,byte ptr [eax]
    lea eax,szLocalTempBuffer
    mov byte ptr [eax],cl
    mov byte ptr [eax+1],00H
    Invoke ReplaceMacros,addr szTargetBuffer,addr szDRIVE,addr szLocalTempBuffer       ;Drive that is redirected to ShareAddress (D  db  C  db  etc)
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ERRORCODE ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke wsprintf,addr szLocalTempBuffer,addr szIntFormat,dErrorCode
    ;Invoke dwtoa, dErrorCode,addr szLocalTempBuffer
    Invoke ReplaceMacros,addr szTargetBuffer,addr szERRORCODE,addr szLocalTempBuffer   ;Last Error Code Returned by  GetLastError
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ ERRORTEXT ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke FormatMessage, FORMAT_MESSAGE_FROM_SYSTEM, 0, dErrorCode, 0, addr szLocalTempBuffer,MAX_PATH, 0
    .if szLocalTempBuffer
        Invoke ReplaceMacros,addr szTargetBuffer,addr szERRORTEXT,addr szLocalTempBuffer   ;Last Error Text Returned by  GetLastError
    .endif
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ PROGRAMDIR ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke lstrcpy,addr szLocalTempBuffer ,addr szProgramPath
    Invoke lstrlen, addr szLocalTempBuffer
    lea ecx,szLocalTempBuffer
    @@:
    dec eax                                           ;Decrement the string
    cmp byte ptr [ecx+eax],'\'
    je @F
        jmp @B
    @@:
    mov byte ptr [ecx+eax],00H
    Invoke ReplaceMacros,addr szTargetBuffer,addr szPROGRAMDIR,addr szLocalTempBuffer  ;Path of the Program Working directory
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ PROGRAMFILES ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke RegOpenKeyEx,HKEY_LOCAL_MACHINE,addr szCurrentVersion,NULL,KEY_READ,addr hReg
    .if eax==ERROR_SUCCESS
        mov dSize,MAX_PATH
        Invoke RegQueryValueEx,hReg,addr szProgramFilesDir,NULL,NULL,addr szLocalTempBuffer,addr dSize
        Invoke RegCloseKey,hReg
        Invoke lstrlen, addr szLocalTempBuffer
        lea ecx,szLocalTempBuffer
        dec eax                                           ;Decrement the string
        cmp byte ptr [ecx+eax],'\'
        jne @F
            ;inc eax
            ;mov byte ptr [ecx+eax],'\'
            ;inc eax
            mov byte ptr [ecx+eax],00H
        @@:
        Invoke ReplaceMacros,addr szTargetBuffer,addr szPROGRAMFILES,addr szLocalTempBuffer;Path of the Program Files Directory
    .endif
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ PROGRAMNAME ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke ReplaceMacros,addr szTargetBuffer,addr szPROGRAMNAME,addr szProgramName  ;Name of the current executable file  usually = MapDrive
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ SHARE ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    ;Invoke ReplaceMacros,addr szTargetBuffer,addr szSHARE,addr szShareAddress       ;Full Share Address
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ SYSTEMDIR ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke GetSystemDirectory,addr szLocalTempBuffer,MAX_PATH
    Invoke lstrlen, addr szLocalTempBuffer
    lea ecx,szLocalTempBuffer
    dec eax                                           ;Decrement the string
    cmp byte ptr [ecx+eax],'\'
    jne @F
        ;inc eax
        ;mov byte ptr [ecx+eax],'\'
        ;inc eax
        mov byte ptr [ecx+eax],00H
    @@:
    Invoke ReplaceMacros,addr szTargetBuffer,addr szSYSTEMDIR,addr szLocalTempBuffer     ;Path of the Windows system directory
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ SYSTEMDRIVE ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    lea ecx,szLocalTempBuffer
    add ecx,3
    mov byte ptr [ecx+eax],00H
    Invoke ReplaceMacros,addr szTargetBuffer,addr szSYSTEMDRIVE,addr szLocalTempBuffer   ;Default Drive Directory usually = C:\
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ TEMPDIR ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke GetTempPath,MAX_PATH,addr szLocalTempBuffer
    lea ecx,szLocalTempBuffer
    dec eax                                           ;Decrement the string
    cmp byte ptr [ecx+eax],'\'
    jne @F
        ;inc eax
        ;mov byte ptr [ecx+eax],'\'
        ;inc eax
        mov byte ptr [ecx+eax],00H
    @@:
    Invoke ReplaceMacros,addr szTargetBuffer,addr szTEMPDIR,addr szLocalTempBuffer       ;path of the directory designated for temporary files
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ USERNAME ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    mov dSize,MAX_PATH
    Invoke GetUserName,addr szLocalTempBuffer,addr dSize
    Invoke ReplaceMacros,addr szTargetBuffer,addr szUSERNAME,addr szLocalTempBuffer      ;Current user's Windows user ID
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ COMPUTERNAME ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    mov dSize,MAX_PATH
    Invoke GetComputerName,addr szLocalTempBuffer,addr dSize
    Invoke ReplaceMacros,addr szTargetBuffer,addr szCOMPUTERNAME,addr szLocalTempBuffer  ;Computer name of the current system
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ WINDIR ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke GetWindowsDirectory, addr szWindowsDir,MAX_PATH
    Invoke lstrlen, addr szWindowsDir ;No Needed, cos GetWindowsDirectory return on eax, len of Windows directory
    lea ecx,szWindowsDir
    dec eax                                           ;Decrement the string
    cmp byte ptr [ecx+eax],'\'
    jne @F
        ;inc eax
        ;mov byte ptr [ecx+eax],'\'
        ;inc eax
        mov byte ptr [ecx+eax],00H
    @@:
    ;Invoke lstrcpy,addr szWindowsDir,addr szWindowsPath
    Invoke ReplaceMacros,addr szTargetBuffer,addr szWINDIR ,addr szWindowsDir           ;Path of the Windows directory
    ;¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬ COUNTER ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Invoke wsprintfA,addr szLocalTempBuffer, addr szIntFormat,dSnapShotCounter
    Invoke ReplaceMacros,addr szTargetBuffer,addr szCOUNTER,addr szLocalTempBuffer
    ;
    Invoke lstrcpy,lpTargetStr,addr szTargetBuffer
    ret
SolveMacros                     endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
;           *** ReplaceSubStr **
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
ReplaceMacros                   proc  lpSourceStr:DWORD, lpPattern:DWORD, lpTarget:DWORD
                          ;
                          LOCAL pLen:DWORD
                          LOCAL dTextPos:DWORD
                          LOCAL szTempCopy[MAX_PATH]:BYTE
                          LOCAL szLocalTempBuffer[MAX_PATH]:BYTE

    Invoke lstrcmp,lpSourceStr,lpPattern
    .if eax==0
        Invoke lstrcpy,lpSourceStr,lpTarget
        ret
    .endif
    Invoke StrLen,lpPattern
    mov pLen, eax                       ; Pattern length
    @@:
      Invoke lstrcpy,addr szTempCopy,lpSourceStr ;Source to Temp
      Invoke InString,1,addr szTempCopy,lpPattern  ;0=No found, -1=same length or longer -2=out of range, >0 = TextPos
      cmp eax,1                         ;Bail if less than 1
      jl  @F
          mov dTextPos,eax              ;it returns the 1 based index of the start of the substring.
          Invoke lstrcpyn, addr szLocalTempBuffer,addr szTempCopy,eax  ;Copy leading char on Buffer
          Invoke lstrcat,addr szLocalTempBuffer,lpTarget           ;Appends lpTarget on Buffer
          mov eax,dTextPos              ;First char position of Pattern
          dec eax                       ;Get 0 based index
          mov ecx,pLen
          add eax,ecx                   ;eax==Position of final Pattern char in lpSourceStr
          lea ecx,szTempCopy            ;Address of lpSourceStr
          add ecx,eax                   ;ecx==Address' Position of final Pattern char in lpSourceStr
          Invoke lstrcat,addr szLocalTempBuffer,ecx  ;Appends Rest of lpTarget on Buffer
          Invoke lstrcpy,lpSourceStr,addr szLocalTempBuffer ;Copy back Buffer to Target
          cmp eax,NULL ;fail???
          je @F
      jmp @B
    @@:
    ;Invoke MessageBox,NULL,lpPattern,addr szTempCopy,MB_OK OR MB_ICONEXCLAMATION
  ret
ReplaceMacros                   endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
;           *** PasteMacro **
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
PasteMacro                      proc  dPopUpSelection :DWORD, hEditHandle :DWORD, dEditID

                          LOCAL wSelStartPos:WORD
                          LOCAL wSelEndPos:WORD
                          LOCAL pMacroText:DWORD
                          LOCAL szTempCopy[MAX_PATH]:BYTE
                          LOCAL szLocalTempBuffer[MAX_PATH]:BYTE

    .if   dPopUpSelection == IDM_01
        mov eax, OFFSET szCOMPUTERNAME  ;Computer name of the current system
    .elseif dPopUpSelection == IDM_02
        mov eax, OFFSET szCOUNTER        ;dSnapShotCounter
    .elseif dPopUpSelection == IDM_03
        mov eax, OFFSET szDATE          ;Current date (in the format YYYYMMDD)
    .elseif dPopUpSelection == IDM_04
        mov eax, OFFSET szDAY           ;Day of the month (1-31)
    .elseif dPopUpSelection == IDM_05
        mov eax, OFFSET szDRIVE         ;Drive that is redirected to ShareAddress (D  db  C  db  etc)
    .elseif dPopUpSelection == IDM_06
        mov eax, OFFSET szERRORCODE     ;Last Error Code Returned by  GetLastError
    .elseif dPopUpSelection == IDM_07
        mov eax, OFFSET szERRORTEXT     ;Last Error Text Returned by  GetLastError
    .elseif dPopUpSelection == IDM_08
        mov eax, OFFSET szHOUR                ;LocalTime Hour 0-24
    .elseif dPopUpSelection == IDM_09
        mov eax, OFFSET szMINUTE          ;LocalTime Hour Minute 00-59
    .elseif dPopUpSelection == IDM_10
        mov eax, OFFSET szMONTH         ;Month of the Year (1-12)
    .elseif dPopUpSelection == IDM_11
        mov eax, OFFSET szPROGRAMDIR    ;Path of the Program Working directory
    .elseif dPopUpSelection == IDM_12
        mov eax, OFFSET szPROGRAMFILES  ;Path of the Program Files Directory
    .elseif dPopUpSelection == IDM_13
        mov eax, OFFSET szPROGRAMNAME   ;Name of the current executable file  usually = MapDrive
    .elseif dPopUpSelection == IDM_14
        mov eax, OFFSET szSECOND          ;LocalTime Second 00-59
    .elseif dPopUpSelection == IDM_15
        mov eax, OFFSET szSHARE         ;Full Share Address
    .elseif dPopUpSelection == IDM_16
        mov eax, OFFSET szSYSTEMDIR     ;Path of the Windows system directory
    .elseif dPopUpSelection == IDM_17
        mov eax, OFFSET szSYSTEMDRIVE   ;Default Root Drive usually = C:\
    .elseif dPopUpSelection == IDM_18
        mov eax, OFFSET szTEMPDIR          ;Default Temp Directory
    .elseif dPopUpSelection == IDM_19
        mov eax, OFFSET szTIME          ;Current time (in the format HHMMSS)
    .elseif dPopUpSelection == IDM_20
        mov eax, OFFSET szUSERNAME      ;Current user's Windows user ID
    .elseif dPopUpSelection == IDM_21
        mov eax, OFFSET szWEEKDAY       ;Days since Sunday (1 – 7)
    .elseif dPopUpSelection == IDM_22
        mov eax, OFFSET szWINDIR        ;Path of the Windows directory
    .elseif dPopUpSelection == IDM_23
        mov eax, OFFSET szYEAR          ;Current year
    .else
        ret
    .endif
    ;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    mov pMacroText,eax
    mov szTempCopy,0
    Invoke SendMessage,hEditHandle,EM_GETSEL,0,0
    mov wSelStartPos,ax
    shr eax,16
    mov wSelEndPos,ax
    Invoke GetDlgItemText,hOwnerPopMenu,dEditID,addr  szLocalTempBuffer,MAX_PATH ;Copy contenst of edit
    xor eax,eax
    mov ax,wSelStartPos
    inc eax
    Invoke GetDlgItemText,hOwnerPopMenu,dEditID,addr  szTempCopy,eax ;Get leading text before caret
    Invoke lstrcat,addr szTempCopy,pMacroText                                    ;Appends lpTarget on Buffer
    xor eax,eax
    mov ax,wSelEndPos
    lea ecx,szLocalTempBuffer
    add eax,ecx
    Invoke lstrcat,addr szTempCopy,eax                                   ;Appends rest of text
    Invoke SetDlgItemText,hOwnerPopMenu,dEditID,addr  szTempCopy
    Invoke StrLen,pMacroText
    xor ecx,ecx
    mov cx,wSelEndPos
    add eax,ecx
    Invoke SendMessage,hEditHandle,EM_SETSEL,eax,eax
  ret
PasteMacro                      endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;SubWndProc
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
SubWndProc                      proc hWnd:DWORD,uMsg:DWORD,wParam :DWORD,lParam :DWORD
                          ;le
                          LOCAL   pt:POINT
                          LOCAL   rc:RECT
    .if uMsg==WM_CHAR
         .if wParam==VK_SPACE
              Invoke GetKeyState,VK_CONTROL
              and ax,0FFFEH
              .if ax
                  Invoke GetFocus
                  mov hFocus,eax
                  Invoke GetWindowRect,hFocus,addr rc
                  Invoke GetDlgCtrlID, hFocus
                  mov dDummy,eax
                  ;
                  Invoke GetCaretPos,addr pt
                  Invoke SetForegroundWindow, hOwnerPopMenu
                  mov eax,rc.left
                  mov ecx,pt.x
                  add eax,ecx
                  mov pt.x,eax
                  ;
                  mov eax,rc.bottom
                  mov ecx,rc.top
                  sub eax,ecx
                  mov ecx,rc.top
                  add eax,ecx
                  mov ecx,pt.y
                  add eax,ecx
                  mov pt.y,eax
                  ;
                  ;Invoke MessageBeep,MB_ICONHAND
                  Invoke TrackPopupMenu, hMacroPopupMenu, TPM_LEFTALIGN + TPM_RETURNCMD, pt.x, pt.y,NULL, hOwnerPopMenu, NULL
                  .if eax
                        Invoke PasteMacro,eax,hFocus,dDummy  ; PopUpSelection,EditHandle,EditID
                  .endif
                  xor eax,eax
                  ret
              .endif
          .endif
    .endif
    ;

    mov ecx,hWnd
    lea eax,ActiveIDC
    @@:
        cmp dword ptr [eax],ecx
        je @F
        add eax,8
        jmp @B
    @@:
    add eax,4
    mov eax,dword ptr [eax]
    Invoke CallWindowProc,eax,hWnd,uMsg,wParam,lParam
    ret
SubWndProc                      endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;ServiceManager  PROCEDURE
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
ServiceManager                  proc dServiceName:DWORD, dServiceDesc:DWORD,dServicePath:DWORD, bAction:BYTE
                    LOCAL hSCManager:DWORD
                    LOCAL hService:DWORD
                    LOCAL SeviceStatus:SERVICE_STATUS
                    LOCAL dReturn:DWORD
                    LOCAL pTitle:DWORD
                    ;INSTALL         equ 0
                    ;REMOVE          equ 1
                    ;ISINSTALLED     equ 2
                    ;ISRUNNING       equ 3
      ;
      mov dReturn,FALSE
      mov hSCManager,NULL
      mov hService,NULL
      .if (bAction==INSTALL) ;INSTALL,REMOVE,ISINSTALLED,ISRUNNING
            ;********************************************************************
            ;************* open a connection to the SCM *************************
            Invoke OpenSCManager,NULL,NULL,GENERIC_WRITE ;SC_MANAGER_CREATE_SERVICE
            mov hSCManager,eax
            .if eax==NULL ;If the function fails, the return value is NULL
                szText szOpenSCM,"Error 0pening a Connection to the SCM"
                lea eax,szOpenSCM
                mov pTitle,eax
                jmp SendErrorMessage
            .endif
            ;********************************************************************
            ;************* Install the new service ******************************
            Invoke CreateService,hSCManager,dServiceName,dServiceDesc,SERVICE_ALL_ACCESS,SERVICE_WIN32_OWN_PROCESS, \
                  SERVICE_DEMAND_START,SERVICE_ERROR_NORMAL,dServicePath,0,0,0,0,0
            mov hService,eax
            .if eax==NULL ;If the function fails, the return value is NULL
                szText szInstall,"Error Creating the new service"
                lea eax,szInstall
                mov pTitle,eax
                jmp SendErrorMessage
            .endif
            ;********************************************************************
            ;************* clean up *********************************************
            mov dReturn,TRUE
            jmp fin
      .else ;.if (bAction==INSTALL) ;INSTALL,REMOVE,ISINSTALLED,ISRUNNING
            ;********************************************************************
            ;************* Open a connection to the SCM**************************
            Invoke OpenSCManager,NULL,NULL,SC_MANAGER_CREATE_SERVICE ;GENERIC_WRITE
            mov hSCManager,eax
            .if eax==NULL ;If the function fails, the return value is NULL
                ;szText szOpenSCM,"Error 0pening a Connection to the SCM"
                lea eax,szOpenSCM
                mov pTitle,eax
                jmp SendErrorMessage
            .endif
            ;********************************************************************
            ;************* Get the service's handle
            Invoke OpenService,hSCManager,dServiceName,SERVICE_ALL_ACCESS OR DELETE
            mov hService,eax
            .if eax==NULL ;If the function fails, the return value is NULL
                .if bAction==ISINSTALLED
                    mov dReturn,FALSE
                    jmp fin
                .endif
                szText szOpenService,"Error 0pening a Service of the SCM"
                lea eax,szOpenService
                mov pTitle,eax
                jmp SendErrorMessage
            .endif
            .if bAction==ISINSTALLED
                mov dReturn,TRUE
                jmp fin
            .endif
            ;********************************************************************
            ;************* Query Service Status**********************************
            Invoke QueryServiceStatus,hService,addr SeviceStatus
            .if eax==0 ;If the function succeeds, the return value is nonzero
                szText szQueryServiceStatus,"Error at Query Service Status of the SCM"
                lea eax,szQueryServiceStatus
                mov pTitle,eax
                jmp SendErrorMessage
            .endif
            ;********************************************************************
            ;************* Stop the service if necessary ************************
            .if SeviceStatus.dwCurrentState != SERVICE_STOPPED
                .if bAction==ISRUNNING
                    mov dReturn,TRUE
                    jmp fin
                .endif
                Invoke ControlService,hService,SERVICE_CONTROL_STOP,addr SeviceStatus
                .if eax==0 ;If the function fails, the return value is zero.
                    szText szStoppingService,"Error Stopping Service"
                    lea eax,szStoppingService
                    mov pTitle,eax
                    jmp SendErrorMessage
                .endif
            .else
                .if bAction==ISRUNNING
                    mov dReturn,FALSE
                    jmp fin
                .endif
            .endif
            Invoke Sleep,1000
            ;********************************************************************
            ;************* Remove the service ***********************************
            Invoke DeleteService,hService
            .if eax==0 ;If the function fails, the return value is zero.
                szText szDeleteService,"Error Deleting Service"
                lea eax,szDeleteService
                mov pTitle,eax
                jmp SendErrorMessage
            .endif
            ;********************************************************************
            ;************* clean up *********************************************
            mov dReturn,TRUE
            jmp fin
      .endif   ;.if (bAction==INSTALL) ;INSTALL,REMOVE,ISINSTALLED,ISRUNNING
      SendErrorMessage:
          Invoke ErrorManager,pTitle

      fin:
          .if hService!=NULL
              Invoke CloseServiceHandle,hService
          .endif
          .if hSCManager!=NULL
              Invoke CloseServiceHandle,hSCManager
          .endif
          mov eax,dReturn
      ret
ServiceManager                  endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;ErrorManager  PROCEDURE
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
ErrorManager                    proc pTitle:DWORD
                    LOCAL szErrorBuff[MAX_PATH]:BYTE
                    LOCAL szErrorLocal[MAX_PATH]:BYTE
                    LOCAL dLastErrorCode:DWORD
                    szText szErrorText,"Error "
    ;
    Invoke GetLastError
    mov dLastErrorCode,eax
    Invoke FormatMessage, FORMAT_MESSAGE_FROM_SYSTEM, 0, dLastErrorCode, 0, addr szErrorBuff,MAX_PATH, 0
    Invoke wsprintf,addr szErrorLocal,addr szIntFormat,dLastErrorCode
    ;Invoke dwtoa, dLastErrorCode,addr szErrorLocal
    Invoke lstrlen,addr szErrorLocal
    lea ecx,szErrorLocal
    mov byte ptr [eax+ecx],':'
    mov byte ptr [eax+ecx+1],' '
    mov byte ptr [eax+ecx+2],0
    Invoke lstrcat,addr szErrorLocal,addr szErrorBuff
    Invoke lstrcpy,addr szErrorBuff,addr szErrorText
    Invoke lstrcat,addr szErrorBuff,addr szErrorLocal
    Invoke MessageBox,NULL,addr szErrorBuff,pTitle,MB_OK+MB_ICONEXCLAMATION
    ;
    ret
    ;
ErrorManager                    endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;Paint Procedure
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
Paint_Proc                      proc hWin:DWORD, xPos:DWORD, yPos:DWORD, wPos:DWORD, hPos:DWORD
                        ;
                        LOCAL hDC   :DWORD
                        LOCAL hOld  :DWORD
                        LOCAL memDC :DWORD
                        LOCAL ps    :PAINTSTRUCT
      ;
      Invoke BeginPaint,hWin,ADDR ps
      mov hDC, eax
      Invoke CreateCompatibleDC,hDC
      mov memDC, eax
      Invoke SelectObject,memDC,hBmp
      mov hOld, eax
      Invoke BitBlt,hDC,xPos,yPos,wPos,hPos,memDC,0,0,SRCCOPY
      Invoke SelectObject,hDC,hOld
      Invoke DeleteDC,memDC
      Invoke EndPaint,hWin,ADDR ps
      ;
      ret
      ;
Paint_Proc                      endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;HyperLink Relative
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
OPTION PROLOGUE:NONE
OPTION EPILOGUE:NONE
_HyperlinkParentProc            proc hWin:DWORD,uMsg:DWORD,wParam:DWORD,lParam:DWORD
      thisPROC_hWnd TEXTEQU <[esp+4*1]>
      thisPROC_uMsg TEXTEQU <[esp+4*2]>
      thisPROC_wParam TEXTEQU <[esp+4*3]>
      thisPROC_lParam TEXTEQU <[esp+4*4]>

      invoke  GetProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_PROC
      mov ecx, thisPROC_uMsg
      pop edx
      push  eax
      cmp ecx, WM_CTLCOLORSTATIC
      push  edx ; return address
      je  __WM_CTLCOLORSTATIC
      cmp ecx, WM_DESTROY
      je  __WM_DESTROY
      jmp CallWindowProc

      ; another dword on the stack
      thisPROC_old  TEXTEQU <[4][esp+4*0]>
      thisPROC_hWnd TEXTEQU <[4][esp+4*1]>
      thisPROC_uMsg TEXTEQU <[4][esp+4*2]>
      thisPROC_wParam TEXTEQU <[4][esp+4*3]>
      thisPROC_lParam TEXTEQU <[4][esp+4*4]>

      __WM_DESTROY:
      invoke  SetWindowLong, thisPROC_hWnd[8], GWL_WNDPROC, eax
      invoke  RemoveProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_PROC
      _x: jmp CallWindowProc


      __WM_CTLCOLORSTATIC:
      ; why would another control have this property?
      invoke  GetProp, thisPROC_lParam[4],addr PROP_ORIGINAL_PROC
      test  eax, eax
      je  _x
      invoke  CallWindowProc, thisPROC_old[16], thisPROC_hWnd[12], thisPROC_uMsg[8], thisPROC_wParam[4], thisPROC_lParam
      mov [esp+4], eax
      invoke  SetTextColor, thisPROC_wParam[4],00C40000H ;CONST_RGB ;(0,0,192)
      mov eax, [esp+4]
      retn 20
_HyperlinkParentProc            endp
OPTION PROLOGUE:PROLOGUEDEF
OPTION EPILOGUE:EPILOGUEDEF
;
align 4
OPTION PROLOGUE:NONE
OPTION EPILOGUE:NONE
_HyperlinkProc                  proc hWin:DWORD,uMsg:DWORD,wParam:DWORD,lParam:DWORD
      thisPROC_hWnd TEXTEQU <[esp+4*1]>
      thisPROC_uMsg TEXTEQU <[esp+4*2]>
      thisPROC_wParam TEXTEQU <[esp+4*3]>
      thisPROC_lParam TEXTEQU <[esp+4*4]>

      invoke  GetProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_PROC
      mov ecx, thisPROC_uMsg
      pop edx
      push  eax
      cmp ecx, WM_MOUSEMOVE
      push  edx ; return address
      je  __WM_MOUSEMOVE
      cmp ecx, WM_SETCURSOR
      je  __WM_SETCURSOR
      cmp ecx, WM_DESTROY
      je  __WM_DESTROY
      jmp CallWindowProc

      ; another dword on the stack
      thisPROC_old  TEXTEQU <[4][esp+4*0]>
      thisPROC_hWnd TEXTEQU <[4][esp+4*1]>
      thisPROC_uMsg TEXTEQU <[4][esp+4*2]>
      thisPROC_wParam TEXTEQU <[4][esp+4*3]>
      thisPROC_lParam TEXTEQU <[4][esp+4*4]>

      __WM_DESTROY:
      invoke  SetWindowLong, thisPROC_hWnd[8], GWL_WNDPROC, eax
      invoke  RemoveProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_PROC

      invoke  GetProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_FONT
      invoke  SendMessage, thisPROC_hWnd[12], WM_SETFONT, eax, 0
      invoke  RemoveProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_FONT

      invoke  GetProp, thisPROC_hWnd[4],addr PROP_UNDERLINE_FONT
      invoke  DeleteObject, eax
      invoke  RemoveProp, thisPROC_hWnd[4],addr PROP_UNDERLINE_FONT
      jmp CallWindowProc


      __WM_MOUSEMOVE:
      invoke  GetCapture
      cmp eax, thisPROC_hWnd
      jne _start_capture

      sub esp, SIZEOF RECT
      invoke  GetWindowRect, thisPROC_hWnd[4 + SIZEOF RECT], esp
      movsx eax, WORD PTR thisPROC_lParam[2 + SIZEOF RECT]
      movsx edx, WORD PTR thisPROC_lParam[0 + SIZEOF RECT]
      push  eax ; y
      push  edx ; x
      invoke  ClientToScreen, thisPROC_hWnd[4 + SIZEOF RECT + SIZEOF POINT], esp
      ; (x, y) already on stack
      push  esp
      add DWORD PTR [esp], SIZEOF POINT
      call  PtInRect

      test  eax, eax
      lea esp, [esp + SIZEOF RECT]
      jne _x
      invoke  GetProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_FONT
      invoke  SendMessage, thisPROC_hWnd[12], WM_SETFONT, eax, FALSE
      invoke  InvalidateRect, thisPROC_hWnd[8], NULL, FALSE
      invoke  ReleaseCapture
      _x: jmp CallWindowProc

      _start_capture:
      invoke  GetProp, thisPROC_hWnd[4],addr PROP_UNDERLINE_FONT
      invoke  SendMessage, thisPROC_hWnd[12], WM_SETFONT, eax, FALSE
      invoke  InvalidateRect, thisPROC_hWnd[8], NULL, FALSE
      invoke  SetCapture, thisPROC_hWnd
      jmp CallWindowProc


      __WM_SETCURSOR:
      ; Since IDC_HAND is not available on all operating systems,
      ; we will load the arrow cursor if IDC_HAND is not present.
      mov edx, IDC_HAND
      @@: invoke  LoadCursor, NULL, edx
      ;@@: invoke  LoadCursor,hInstance,550
      test  eax, eax
      mov edx, IDC_ARROW
      je  @B
      invoke  SetCursor, eax
      mov eax, TRUE
      ret 5*4
_HyperlinkProc                  endp
OPTION PROLOGUE:PROLOGUEDEF
OPTION EPILOGUE:EPILOGUEDEF
align 4
ConvertStaticToHyperlink        proc hwndParent:DWORD, uiCtlId:DWORD

      Invoke GetDlgItem, hwndParent, uiCtlId
      mov uiCtlId, eax ; need handle, not Id

      ; Subclass the parent so we can color the controls as we desire
      cmp hwndParent, NULL
      je  @F
          invoke  GetWindowLong, hwndParent, GWL_WNDPROC
      ;
      cmp eax, OFFSET _HyperlinkParentProc
      je  @F
          Invoke  SetProp, hwndParent,addr PROP_ORIGINAL_PROC, eax
          Invoke  SetWindowLong, hwndParent, GWL_WNDPROC, OFFSET _HyperlinkParentProc
      @@:
      ; Make sure the control will send notifications
      ;
      invoke  GetWindowLong, uiCtlId, GWL_STYLE
      or  eax, SS_NOTIFY
      invoke  SetWindowLong, uiCtlId, GWL_STYLE, eax

      ; Subclass the existing control

      invoke  GetWindowLong, uiCtlId, GWL_WNDPROC
      invoke  SetProp, uiCtlId,addr PROP_ORIGINAL_PROC, eax
      invoke  SetWindowLong, uiCtlId, GWL_WNDPROC, OFFSET _HyperlinkProc

      ; Create an updated font by adding an underline

      invoke  SendMessage, uiCtlId, WM_GETFONT, 0, 0
      push  eax
      invoke  SetProp, uiCtlId,addr PROP_ORIGINAL_FONT, eax
      pop edx
      sub esp, SIZEOF LOGFONT
      invoke  GetObject, edx, SIZEOF LOGFONT, esp
      mov [esp].LOGFONT.lfUnderline, TRUE

      invoke  CreateFontIndirect, esp
      invoke  SetProp, uiCtlId,addr PROP_UNDERLINE_FONT, eax
      add esp, SIZEOF LOGFONT

      mov eax, TRUE
      ret
ConvertStaticToHyperlink        endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;GETFILENAME  PROCEDURE
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
GetFileName                     proc hParent:DWORD,lpTitle:DWORD,lpFilter:DWORD

      mov ofn.lStructSize,        sizeof OPENFILENAME
      m2m ofn.hWndOwner,          hParent
      m2m ofn.hInstance,          hInstance
      m2m ofn.lpstrFilter,        lpFilter
      m2m ofn.lpstrFile,          offset szBuffer
      mov ofn.nMaxFile,           sizeof szBuffer
      m2m ofn.lpstrTitle,         lpTitle
      mov ofn.Flags,              OFN_EXPLORER or \
                                  OFN_LONGNAMES or OFN_NOCHANGEDIR
      ;or OFN_FILEMUSTEXIST
      Invoke GetOpenFileNameA,ADDR ofn

    ret

GetFileName                     endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;FillBuffer  PROCEDURE
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
FillBuffer                      proc lpBuffer:DWORD,lenBuffer:DWORD,TheChar:BYTE

      push edi

      mov edi, lpBuffer   ; address of buffer
      mov ecx, lenBuffer  ; buffer length
      mov  al, TheChar    ; load al with character
      rep stosb           ; write character to buffer until ecx = 0

      pop edi

      ret

FillBuffer                      endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;Varias
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
StrLen                          proc item:DWORD

  ; -------------------------------------------------------------
  ; This procedure has been adapted from an algorithm written by
  ; Agner Fog. It has the unusual characteristic of reading up to
  ; three bytes past the end of the buffer as it uses DWORD size
  ; reads. It is measurably faster than a classic byte scanner on
  ; large linear reads and has its place where linear read speeds
  ; are important.
  ; -------------------------------------------------------------

    push    ebx
    mov     eax,item               ; get pointer to string
    lea     edx,[eax+3]            ; pointer+3 used in the end
  @@:
    mov     ebx,[eax]              ; read first 4 bytes
    add     eax,4                  ; increment pointer
    lea     ecx,[ebx-01010101h]    ; subtract 1 from each byte
    not     ebx                    ; invert all bytes
    and     ecx,ebx                ; and these two
    and     ecx,80808080h
    jz      @B                     ; no zero bytes, continue loop
    test    ecx,00008080h          ; test first two bytes
    jnz     @F
    shr     ecx,16                 ; not in the first 2 bytes
    add     eax,2
  @@:
    shl     cl,1                   ; use carry flag to avoid branch
    sbb     eax,edx                ; compute length
    pop     ebx

    ret

StrLen                          endp
align 4
InString                        proc startpos:DWORD,lpSource:DWORD,lpPattern:DWORD

  ; ------------------------------------------------------------------
  ; InString searches for a substring in a larger string and if it is
  ; found, it returns its position in eax.
  ;
  ; It uses a one (1) based character index (1st character is 1,
  ; 2nd is 2 etc...) for both the "StartPos" parameter and the returned
  ; character position.
  ;
  ; Return Values.
  ; If the function succeeds, it returns the 1 based index of the start
  ; of the substring.
  ;  0 = no match found
  ; -1 = substring same length or longer than main string
  ; -2 = "StartPos" parameter out of range (less than 1 or longer than
  ; main string)
  ; ------------------------------------------------------------------

    LOCAL sLen:DWORD
    LOCAL pLen:DWORD

    push ebx
    push esi
    push edi

    invoke StrLen,lpSource
    mov sLen, eax           ; source length
    invoke StrLen,lpPattern
    mov pLen, eax           ; pattern length

    cmp startpos, 1
    jge @F
    mov eax, -2
    jmp isOut               ; exit if startpos not 1 or greater
  @@:

    dec startpos            ; correct from 1 to 0 based index

    cmp  eax, sLen
    jl @F
    mov eax, -1
    jmp isOut               ; exit if pattern longer than source
  @@:

    sub sLen, eax           ; don't read past string end
    inc sLen

    mov ecx, sLen
    cmp ecx, startpos
    jg @F
    mov eax, -2
    jmp isOut               ; exit if startpos is past end
  @@:

  ; ----------------
  ; setup loop code
  ; ----------------
    mov esi, lpSource
    mov edi, lpPattern
    mov al, [edi]           ; get 1st char in pattern

    add esi, ecx            ; add source length
    neg ecx                 ; invert sign
    add ecx, startpos       ; add starting offset

    jmp Scan_Loop

    align 16

  ; @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

  Pre_Scan:
    inc ecx                 ; start on next byte

  Scan_Loop:
    cmp al, [esi+ecx]       ; scan for 1st byte of pattern
    je Pre_Match            ; test if it matches
    inc ecx
    js Scan_Loop            ; exit on sign inversion

    jmp No_Match

  Pre_Match:
    lea ebx, [esi+ecx]      ; put current scan address in EBX
    mov edx, pLen           ; put pattern length into EDX

  Test_Match:
    mov ah, [ebx+edx-1]     ; load last byte of pattern length in main string
    cmp ah, [edi+edx-1]     ; compare it with last byte in pattern
    jne Pre_Scan            ; jump back on mismatch
    dec edx
    jnz Test_Match          ; 0 = match, fall through on match

  ; @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

  Match:
    add ecx, sLen
    mov eax, ecx
    inc eax
    jmp isOut

  No_Match:
    xor eax, eax

  isOut:
    pop edi
    pop esi
    pop ebx

    ret

InString                        endp
align 4
GetCL                           proc ArgNum:DWORD, ItemBuffer:DWORD

  ; -------------------------------------------------
  ; arguments returned in "ItemBuffer"
  ;
  ; arg 0 = program name
  ; arg 1 = 1st arg
  ; arg 2 = 2nd arg etc....
  ; -------------------------------------------------
  ; Return values in eax
  ;
  ; 1 = successful operation
  ; 2 = no argument exists at specified arg number
  ; 3 = non matching quotation marks
  ; 4 = empty quotation marks
  ; -------------------------------------------------

    LOCAL lpCmdLine      :DWORD
    LOCAL cmdBuffer[192] :BYTE
    LOCAL tmpBuffer[192] :BYTE

    push esi
    push edi

    invoke GetCommandLine
    mov lpCmdLine, eax        ; address command line

  ; -------------------------------------------------
  ; count quotation marks to see if pairs are matched
  ; -------------------------------------------------
    xor ecx, ecx            ; zero ecx & use as counter
    mov esi, lpCmdLine

    @@:
      lodsb
      cmp al, 0
      je @F
      cmp al, 34            ; [ " ] character
      jne @B
      inc ecx               ; increment counter
      jmp @B
    @@:

    push ecx                ; save count

    shr ecx, 1              ; integer divide ecx by 2
    shl ecx, 1              ; multiply ecx by 2 to get dividend

    pop eax                 ; put count in eax
    cmp eax, ecx            ; check if they are the same
    je @F
      pop edi
      pop esi
      mov eax, 3            ; return 3 in eax = non matching quotation marks
      ret
    @@:

  ; ------------------------
  ; replace tabs with spaces
  ; ------------------------
    mov esi, lpCmdLine
    lea edi, cmdBuffer

    @@:
      lodsb
      cmp al, 0
      je rtOut
      cmp al, 9     ; tab
      jne rtIn
      mov al, 32
    rtIn:
      stosb
      jmp @B
    rtOut:
      stosb         ; write last byte

  ; -----------------------------------------------------------
  ; substitute spaces in quoted text with replacement character
  ; -----------------------------------------------------------
    lea eax, cmdBuffer
    mov esi, eax
    mov edi, eax

    subSt:
      lodsb
      cmp al, 0
      jne @F
      jmp subOut
    @@:
      cmp al, 34
      jne subNxt
      stosb
      jmp subSl     ; goto subloop
    subNxt:
      stosb
      jmp subSt

    subSl:
      lodsb
      cmp al, 32    ; space
      jne @F
        mov al, 254 ; substitute character
      @@:
      cmp al, 34
      jne @F
        stosb
        jmp subSt
      @@:
      stosb
      jmp subSl

    subOut:
      stosb         ; write last byte

  ; ----------------------------------------------------
  ; the following code determines the correct arg number
  ; and writes the arg into the destination buffer
  ; ----------------------------------------------------
    lea eax, cmdBuffer
    mov esi, eax
    lea edi, tmpBuffer

    mov ecx, 0          ; use ecx as counter

  ; ---------------------------
  ; strip leading spaces if any
  ; ---------------------------
    @@:
      lodsb
      cmp al, 32
      je @B

    l2St:
      cmp ecx, ArgNum     ; the number of the required cmdline arg
      je clSubLp2
      lodsb
      cmp al, 0
      je cl2Out
      cmp al, 32
      jne cl2Ovr           ; if not space

    @@:
      lodsb
      cmp al, 32          ; catch consecutive spaces
      je @B

      inc ecx             ; increment arg count
      cmp al, 0
      je cl2Out

    cl2Ovr:
      jmp l2St

    clSubLp2:
      stosb
    @@:
      lodsb
      cmp al, 32
      je cl2Out
      cmp al, 0
      je cl2Out
      stosb
      jmp @B

    cl2Out:
      mov al, 0
      stosb

  ; ------------------------------
  ; exit if arg number not reached
  ; ------------------------------
    .if ecx < ArgNum
      mov edi, ItemBuffer
      mov al, 0
      stosb
      mov eax, 2  ; return value of 2 means arg did not exist
      pop edi
      pop esi
      ret
    .endif

  ; -------------------------------------------------------------
  ; remove quotation marks and replace the substitution character
  ; -------------------------------------------------------------
    lea eax, tmpBuffer
    mov esi, eax
    mov edi, ItemBuffer

    rqStart:
      lodsb
      cmp al, 0
      je rqOut
      cmp al, 34    ; dont write [ " ] mark
      je rqStart
      cmp al, 254
      jne @F
      mov al, 32    ; substitute space
    @@:
      stosb
      jmp rqStart

  rqOut:
      stosb         ; write zero terminator

  ; ------------------
  ; handle empty quote
  ; ------------------
    mov esi, ItemBuffer
    lodsb
    cmp al, 0
    jne @F
    pop edi
    pop esi
    mov eax, 4  ; return value for empty quote
    ret
  @@:

    mov eax, 1  ; return value success

    pop edi
    pop esi

    ret

GetCL                           endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
a2dw                            proc szString:DWORD


  ; ----------------------------------------
  ; Convert decimal string into dword value
  ; return value in eax
  ; ----------------------------------------

    push esi
    push edi

    xor eax, eax
    mov esi, [szString]
    xor ecx, ecx
    xor edx, edx
    mov al, [esi]
    inc esi
    cmp al, 2D
    jne proceed
    mov al, byte ptr [esi]
    not edx
    inc esi
    jmp proceed

  @@:
    sub al, 30h
    lea ecx, dword ptr [ecx+4*ecx]
    lea ecx, dword ptr [eax+2*ecx]
    mov al, byte ptr [esi]
    inc esi

  proceed:
    or al, al
    jne @B
    lea eax, dword ptr [edx+ecx]
    xor eax, edx

    pop edi
    pop esi

    ret


a2dw                            endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;FindApplicationProcessID  PROCEDURE
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
FindApplicationProcessID        proc pszApplicationPath:DWORD,dWhatToDoIfFound:DWORD,pOutWinHandle:DWORD ;dWhatToDoIfFound: Update OutWinHandle=-1, Hide=0, Unhide=1, Kill=2
                                LOCAL hWin:DWORD
                                LOCAL bLoop:BOOL
                                LOCAL dThreadID:HANDLE
                                LOCAL dReturnProccessID:HANDLE
                                LOCAL pe32:PROCESSENTRY32
                                LOCAL hProcessesSnap:HANDLE
                                ;
                                LOCAL Sh_st_info :STARTUPINFO
                                LOCAL Sh_sfi     :SHFILEINFO
                                ;
                                ;szText  szExtentionDotExe,".exe"

                                LOCAL pFullPathLocal:DWORD
                                LOCAL dlParam[3]:DWORD
        ;
        Invoke FillBuffer,addr szFullPathLocal,sizeof szFullPathLocal,0
        mov dReturnProccessID,NULL
        mov pFullPathLocal,0
        ;
        Invoke SearchPath,NULL,pszApplicationPath,addr szExtentionDotExe,sizeof szFullPathLocal,addr szFullPathLocal,addr pFullPathLocal
        ;
        ;
        ;You should call this function from a background thread. Failure to do so could cause the UI to stop responding.
        Invoke  SHGetFileInfo,addr szFullPathLocal,0,addr Sh_sfi,sizeof Sh_sfi,SHGFI_EXETYPE
                ;Return:
                ;0                                              [Nonexecutable file or an error condition]
                ;LOWORD = NE or PE and HIWORD = Windows version [Windows application]
                ;LOWORD = MZ and HIWORD = 0                     [MS-DOS .exe or .com file]
                ;LOWORD = PE and HIWORD = 0                     [Console application or .bat file]
        .if eax==0
            jmp ExitFindApplicationProcessID
        .endif
        mov ecx,eax
        shr ecx,16
        .if cx==0
            ;ax==4550H -> WIN32_CON = PE [Console application or .bat file]: 4550H from 45H=69dec, 50H=80dec, 69dec ASCII = E and 80dec ASCII = P
            ;ax==5A4DH ->                [MS-DOS .exe or .com file]:                                          90dec ASCII = Z and 77dec ASCII = M
            mov     dword ptr [Sh_st_info.cb],sizeof STARTUPINFO
            Invoke  GetStartupInfoA, addr Sh_st_info
            mov ecx,dword ptr [Sh_st_info.lpTitle]
            mov pszApplicationPath,ecx
            Invoke FindWindow,NULL,pszApplicationPath
            .if eax!=0
                mov hWin,eax
                Invoke GetWindowThreadProcessId,hWin,addr dReturnProccessID
                .if dWhatToDoIfFound==-1
                    mov edx,pOutWinHandle
                    mov eax,hWin
                    mov dword ptr [edx],eax
                .elseif dWhatToDoIfFound==0 ;dWhatToDoIfFound Hide=0 Unhide=1 Kill=2
                    Invoke ForceWindowToTop,hWin
                    Invoke ShowWindow,hWin,SW_HIDE ;SW_SHOW
                    ;mov dReturnProccessID,eax
                .elseif dWhatToDoIfFound==1
                    Invoke ShowWindow,hWin,SW_SHOW
                    Invoke ForceWindowToTop,hWin
                    ;mov dReturnProccessID,eax
                .else
                    ;Invoke GetWindowThreadProcessId,hWin,addr dReturnProccessID
                    Invoke OpenProcess,PROCESS_ALL_ACCESS,FALSE,dReturnProccessID
                    .if eax!=NULL
                        push eax
                        Invoke TerminateProcess,eax, 0
                        pop eax
                        Invoke CloseHandle,eax
                        ;mov dReturnProccessID,
                    .endif
                .endif
            .endif
            ;
        .else ;.if cx==0
            ;cx=Window Version, ax = NE or PE       [Windows application: WIN32_EXE]
            mov dReturnProccessID,NULL
            ;
            Invoke CreateToolhelp32Snapshot,TH32CS_SNAPPROCESS,0 ;Takes a snapshot of the specified processes, as well as the heaps, modules, and threads used by these processes.
            .if eax == INVALID_HANDLE_VALUE
                xor eax,eax
                jmp ExitFindApplicationProcessID
            .endif
            mov hProcessesSnap,eax
            mov pe32.dwSize,SIZEOF PROCESSENTRY32
            mov bLoop,TRUE
            ;
            Invoke Process32First,hProcessesSnap,addr pe32
            .if eax==FALSE ;Returns TRUE if the first entry of the process list has been copied to the buffer or FALSE otherwise.
                Invoke CloseHandle,hProcessesSnap
                xor eax,eax
                jmp ExitFindApplicationProcessID
            .else
                ;=========================================
                ;TODO:
                ;pe32.szExeFile[MAX_PATH] The name of the executable file for the process.
                ; To retrieve the full path to the executable file, call the Module32First function
                ; and check the szExePath member of the MODULEENTRY32 structure that is returned.
                ; However, if the calling process is a 32-bit process, you must call the QueryFullProcessImageName function
                ; to retrieve the full path of the executable file for a 64-bit process.
                ;=========================================
                .while bLoop
                    Invoke CompareString,LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr pe32.szExeFile, -1, pFullPathLocal, -1
                    .if eax==2
                        ;
                        mov eax,pe32.th32ProcessID
                        mov dReturnProccessID,eax
                        Invoke GetCurrentProcess
                        .if eax==dReturnProccessID
                            jmp NextProcessID
                        .endif
                        ;
                        .if dWhatToDoIfFound == 2 ;dWhatToDoIfFound Hide=0 Unhide=1 Kill=2
                            Invoke OpenProcess,PROCESS_TERMINATE,FALSE,dReturnProccessID ;PROCESS_TERMINATE, PROCESS_ALL_ACCESS
                            .if eax==NULL
                                Invoke OpenProcess,PROCESS_ALL_ACCESS,FALSE,dReturnProccessID ;PROCESS_TERMINATE, PROCESS_ALL_ACCESS
                            .endif
                            .if eax==NULL
                                mov dReturnProccessID,eax
                            .else
                                push eax
                                Invoke TerminateProcess,eax,0
                                .if eax==NULL ;If the function succeeds, the return value is nonzero.
                                    mov dReturnProccessID,eax
                                .endif
                                pop eax
                                Invoke CloseHandle,eax
                                ;mov dReturnProccessID,0
                            .endif
                        .else
                            ;Here we are in the Application Process
                            Invoke FindTopWindowFromProcessID,dReturnProccessID,dWhatToDoIfFound
                            .if pOutWinHandle!=NULL
                                mov ecx,pOutWinHandle
                                mov dword ptr [ecx],eax
                                Invoke IsWindow,eax
                                .if eax
                                    Invoke MessageBoxA,NULL,addr szAtom,NULL,NULL
                                .endif
                            .endif
                        .endif
                        .break
                    .endif
                    NextProcessID:
                    Invoke Process32Next, hProcessesSnap,addr pe32
                    mov bLoop,eax
                .endw
                Invoke CloseHandle,hProcessesSnap
            .endif
        .endif ;.if ( (ax==4550H) && (cx==0) )
        ;
        mov eax,dReturnProccessID
        ;
    ExitFindApplicationProcessID:
        ;
    ret
FindApplicationProcessID        endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
;                       ForceWindowToTop
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
ForceWindowToTop                proc hWnd:DWORD ;********* PUT This Windows ABOVE it all
                                    LOCAL dwCurrentThread :DWORD
                                    LOCAL dwFGThread :DWORD
                                    LOCAL dwProcessID :DWORD
                                    ;szText  szExtentionDotExe,".exe"
                                ;
        ;https://stackoverflow.com/questions/688337/how-do-i-force-my-app-to-come-to-the-front-and-take-focus
        ;
        xor eax,eax
        Invoke IsWindow,hWnd
        .if eax ;If the window handle does not identify an existing window, the return value is zero.
            Invoke GetCurrentThreadId
            mov dwCurrentThread,eax
            Invoke GetForegroundWindow
            Invoke GetWindowThreadProcessId,eax,NULL ;addr dwProcessID
            mov dwFGThread,eax
            Invoke AttachThreadInput,dwCurrentThread,dwFGThread,TRUE
            ;Invoke GetWindowText,hWnd,addr szBuffer,sizeof szBuffer
            ;Invoke MessageBox,hWnd,addr szBuffer,addr szExtentionDotExe,MB_YESNO+MB_ICONQUESTION
            ;Possible actions you may try to bring the window into focus.
            Invoke SetForegroundWindow,hWnd
            .if eax==0 ;If the window was not brought to the foreground, the return value is zero.
                Invoke Sleep,10
                Invoke SwitchToThisWindow,hWnd,FALSE
            .endif
            ;Invoke SetCapture,hWnd
            Invoke SetFocus,hWnd
            Invoke SetActiveWindow,hWnd
            Invoke EnableWindow,hWnd,TRUE
            ;Invoke ShowWindowAsync,hWnd,SW_HIDE
            Invoke ShowWindowAsync,hWnd,SW_SHOWNORMAL
            ;Invoke ShowWindowAsync,hWnd,SW_SHOWNORMAL
            ;
            Invoke AttachThreadInput,dwCurrentThread,dwFGThread,FALSE
            ;
            Invoke GetWindowThreadProcessId,hWnd,addr dwProcessID
            mov eax,dwProcessID
;        .else
;            szText  szOutputDebugString99,"ForceWindowToTop->Invoke IsWindow,hWnd=FALSE"
;            Invoke OutputDebugString, addr szOutputDebugString99
;            xor eax,eax
        .endif
    ret
ForceWindowToTop                endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;EnumWindowsProc  PROCEDURE
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
EnumWindowsProc                 proc hWinInTurn:DWORD,lparam:DWORD
                                    ;
                                    LOCAL dWinMainThreadID:DWORD
                                    LOCAL dWinInTurnThreadID:DWORD
                                    LOCAL dWinInTurnProcessID:DWORD
                                    LOCAL dIsChildWindow:DWORD
                                    ;
                                    LOCAL dThreadID:DWORD
                                    LOCAL dProcessID:DWORD
                                    LOCAL dWhatToDoIfFound:DWORD ;dWhatToDoIfFound Hide=0 Unhide=1 Kill=2
                                    LOCAL pOutWinHandle:DWORD
        ;
        mov edx,lparam
        ;mov eax,dword ptr [edx]
        ;mov dThreadID,eax
        mov eax,dword ptr [edx]
        mov dProcessID,eax
        mov eax,dword ptr [edx+4]
        mov dWhatToDoIfFound,eax
        mov eax,dword ptr [edx+8]
        mov pOutWinHandle,eax
        ;
        Invoke GetProcessMainThreadID,dProcessID
        mov dWinMainThreadID,eax
        ;
        mov dIsChildWindow,FALSE
        Invoke GetWindowThreadProcessId,hWinInTurn,addr dWinInTurnProcessID
        .if eax==NULL
            jmp ExitEnumWindowsProc
        .endif
        mov dWinInTurnThreadID,eax
        mov edx,dProcessID
        .if (dWinInTurnProcessID!=edx)
            ;Filter out every other window where their PID is not dProcessID
            mov dIsChildWindow,TRUE
        .else
            .if dWinMainThreadID
                mov edx,dWinInTurnThreadID
                .if dWinMainThreadID!=edx
                    mov dIsChildWindow,TRUE
                ;.else
                    ;Invoke DwordToAsccii,dWinMainThreadID,addr szOutputDebugString
                    ;Invoke OutputDebugString,addr szOutputDebugString
                    ;Invoke GetWindowText,hWinInTurn,addr szOutputDebugString,sizeof szOutputDebugString
                    ;Invoke MessageBox,hWinInTurn,addr szOutputDebugString,addr szOutputDebugString,NULL
                    ;Invoke OutputDebugString,addr szOutputDebugString
                .endif
            .else
                Invoke GetWindow,hWinInTurn,GW_OWNER
                .if eax!=NULL
                    ;Filter out every windows that has OWNER
                    mov dIsChildWindow,TRUE
                .else
                    Invoke GetParent,hWinInTurn
                    .if eax!=NULL
                        ;Filter out every windows that has PARENT
                        mov dIsChildWindow,TRUE
                    .else
                        ;
                        ;DISCARDED->Invoke IsWindowVisible,hWinInTurn ;Filter does not works if the app is running hidden.
                        ;                  ;If the specified window, its parent window, its parent's parent window, and so forth,
                        ;                  ;have the WS_VISIBLE style, the return value is nonzero. Otherwise, the return value is zero.
                        ;
                        ;TODO: What happen when a Windows has not Title text?
                        Invoke GetWindowText,hWinInTurn,addr szOutputDebugString,sizeof szOutputDebugString
                         .if eax==NULL
                            ;Filter out every windows that has not Text in title.
                            mov dIsChildWindow,TRUE
                        ;.else
                            ;Invoke GetWindowText,hWinInTurn,addr szOutputDebugString,sizeof szOutputDebugString
                            ;Invoke MessageBox,hWinInTurn,addr szOutputDebugString,addr szOutputDebugString,NULL
                            ;Invoke OutputDebugString,addr szOutputDebugString
                        .endif
                    .endif
                .endif
            .endif
        .endif
        .if dIsChildWindow==FALSE
            .if dWhatToDoIfFound==-1 ;dWhatToDoIfFound: Hide=0, Unhide=1, Kill=2
                .if pOutWinHandle!=NULL
                    mov edx,pOutWinHandle
                    mov eax,hWinInTurn
                    mov dword ptr [edx],eax
                .endif
            .elseif dWhatToDoIfFound==0 ;dWhatToDoIfFound: Hide=0, Unhide=1, Kill=2
                Invoke ShowWindow,hWinInTurn,SW_HIDE
            .elseif dWhatToDoIfFound==1
                ;
                mov eax,hWinInTurn
                mov hWinMainInstanceAtom,eax
                .if bWinMainInstanceHotSendKeys==TRUE
                    Invoke SendKey,VK_SHIFT,10,0,TRUE,FALSE
                    Invoke SendKey,VK_CONTROL,10,0,TRUE,FALSE
                    Invoke SendKey,VK_MENU,10,0,TRUE,FALSE;ALT
                    ;
                    Invoke SendKey,VK_S,10,0,TRUE,TRUE ;Invoke SendKey,VK_SNAPSHOT,10,100,TRUE,TRUE
                    ;
                    Invoke SendKey,VK_MENU,10,0,FALSE,TRUE
                    Invoke SendKey,VK_CONTROL,10,0,FALSE,TRUE
                    Invoke SendKey,VK_SHIFT,10,0,FALSE,TRUE
                .endif
                ;
            .endif
        .endif
        ;To continue enumeration, the callback function must return TRUE; to stop enumeration, it must return FALSE.
        ;EnumWindows continues until the last top-level window is enumerated or the callback function returns FALSE.
        mov eax,dIsChildWindow
        ExitEnumWindowsProc:
    ret
EnumWindowsProc                 endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
;                       FindTopWindowFromProcessID
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
FindTopWindowFromProcessID      proc dProcessId:DWORD,dWhatToDoIfFound:DWORD
                                    LOCAL hProcess:HANDLE
                                    LOCAL bLoop:BOOL
                                    LOCAL dOutWinHandle:DWORD
                                    LOCAL dlParam[3]:DWORD
        ;
        mov dOutWinHandle,NULL
        Invoke OpenProcess,PROCESS_ALL_ACCESS,FALSE,dProcessId
        mov hProcess,eax
        .if hProcess!=NULL ;If the function succeeds, the return value is an open handle to the specified process.;
            lea edx,dlParam ;dlParam[3]:DWORD
            mov eax,dProcessId
            mov dword ptr [edx],eax
            mov eax,dWhatToDoIfFound ;dWhatToDoIfFound!=2 ;dWhatToDoIfFound Hide=0 Unhide=1 Kill=2
            mov dword ptr[edx+4],eax
            mov eax,dOutWinHandle
            mov dword ptr [edx+8],eax
            .while bLoop
                ;Enumerates all top-level windows on the screen by passing the handle to each window, in turn,
                ; to an application-defined callback function. EnumWindows continues until
                ; the last top-level window is enumerated or the callback function returns FALSE.
                Invoke EnumWindows,addr EnumWindowsProc,addr dlParam ;pe32.th32ProcessID
                mov bLoop,eax ;Return value Type: BOOL If the function succeeds, the return value is nonzero.
                ;Remarks: The EnumWindows function does not enumerate child windows,
                ; with the exception of a few top-level windows owned by the system that have the WS_CHILD style.
                ;This function is more reliable than calling the GetWindow function in a loop.
                ; An application that calls GetWindow to perform this task risks being caught in an infinite loop or referencing a handle to a window that has been destroyed.
                ;Note  For Windows 8 and later, EnumWindows enumerates only top-level windows of desktop apps.
            .endw
            Invoke CloseHandle,hProcess
        .endif
    mov eax,dOutWinHandle
    ret
FindTopWindowFromProcessID      endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;GetProcessMainThreadID  PROCEDURE
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
GetProcessMainThreadID          proc hProcessID:DWORD
                                LOCAL   dReturnThreadID:HANDLE
                                LOCAL   hThreadsSnap:HANDLE
                                LOCAL   th32:THREADENTRY32
                                LOCAL   bLoop:BOOL
                                LOCAL   hThread:HANDLE
                                LOCAL   afMinCreationTime:FILETIME
                                LOCAL   afCreationTime:FILETIME
                                LOCAL   afExitTime:FILETIME
                                LOCAL   afKernelTime:FILETIME
                                LOCAL   afUserTime:FILETIME
                                ;
        mov dReturnThreadID,NULL
        ;
        Invoke CreateToolhelp32Snapshot,TH32CS_SNAPTHREAD,0 ;Takes a snapshot of the specified processes, as well as the heaps, modules, and threads used by these processes.
        .if eax == INVALID_HANDLE_VALUE
            jmp ExitGetProcessMainThreadID
        .endif
        mov hThreadsSnap,eax
        ;
        mov afMinCreationTime.dwLowDateTime,0FFFFFFFFH
        mov afMinCreationTime.dwHighDateTime,0FFFFFFFFH
        mov th32.dwSize,SIZEOF THREADENTRY32
        mov bLoop,TRUE
        ;
        Invoke Thread32First,hThreadsSnap,addr th32
        .if eax==FALSE ;Returns TRUE if the first entry of the process list has been copied to the buffer or FALSE otherwise.
            Invoke CloseHandle,hThreadsSnap
        .else
            .while bLoop
                ;
                mov eax,hProcessID
                .if th32.th32OwnerProcessID == eax
                    Invoke OpenThread,THREAD_QUERY_INFORMATION,TRUE,th32.th32ThreadID
                    .if eax ;If the function fails, the return value is NULL
                        mov hThread,eax
                        ;
                        Invoke GetThreadTimes,hThread,addr afCreationTime,addr afExitTime,addr afKernelTime,addr afUserTime
                        .if eax ;If the function succeeds, the return value is nonzero.
                            .if afCreationTime.dwLowDateTime!=NULL && afCreationTime.dwHighDateTime!=NULL
                                Invoke CompareFileTime,addr afMinCreationTime,addr afCreationTime
                                ;Return
                                ;-1 = First file time is earlier than second file time.
                                ; 0 = First file time is equal to second file time.
                                ; 1 = First file time is later than second file time.
                                .if eax==1
                                    mov ecx,afCreationTime.dwLowDateTime
                                    mov afMinCreationTime.dwLowDateTime,ecx
                                    mov edx,afCreationTime.dwHighDateTime
                                    mov afMinCreationTime.dwHighDateTime,edx
                                    ;
                                    mov eax,th32.th32ThreadID
                                    mov dReturnThreadID,eax ;let it be main... :)
                                .endif
                            .endif
                        .endif
                        ;
                        Invoke CloseHandle,hThread
                    .endif
                .endif
                ;
                ;
                Invoke Thread32Next,hThreadsSnap,addr th32
                mov bLoop,eax
            .endw
            Invoke CloseHandle,hThreadsSnap
        .endif
        ;
        ExitGetProcessMainThreadID:
        mov eax,dReturnThreadID
    ret
GetProcessMainThreadID          endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;GetProcessMainThreadID  PROCEDURE
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
align 4
SendKey                         proc wVk:DWORD, nSleep:DWORD,bExtendedkey:DWORD,bDown:DWORD,bUp:DWORD
                                LOCAL   input:INPUT
                                ;
        mov input.ddType,INPUT_KEYBOARD
        mov input.ki.wVk,0
        mov input.ki.wScan,0
        mov input.ki.dwFlags,0 ;KEYEVENTF_UNICODE ;OR KEYEVENTF_EXTENDEDKEY ;>>><<<; KEYEVENTF_KEYUP ;KEYEVENTF_UNICODE ;KEYEVENTF_EXTENDEDKEY,KEYEVENTF_KEYUP,KEYEVENTF_SCANCODE
        mov input.ki.time,0
        mov input.ki.dwExtraInfo,0
        ;
        .if bDown
            mov eax,wVk
            mov input.ki.wVk,ax
            Invoke MapVirtualKey,wVk,0
            mov input.ki.wScan,ax
            mov input.ki.dwFlags,KEYEVENTF_SCANCODE ;KEYEVENTF_UNICODE   ;KEYEVENTF_SCANCODE + KEYEVENTF_KEYUP ;
            .if bExtendedkey
                mov input.ki.dwFlags,KEYEVENTF_SCANCODE + KEYEVENTF_EXTENDEDKEY
            .endif
            mov input.ki.time,0
            Invoke  GetMessageExtraInfo
            mov input.ki.dwExtraInfo,eax
            mov input.ddType,INPUT_KEYBOARD
            Invoke SendInput,1,addr input,sizeof INPUT   ;keybd_event
            Invoke Sleep,nSleep
        .endif
        ;
        .if bUp
            mov eax,wVk
            mov input.ki.wVk,ax
            Invoke MapVirtualKey,wVk,0
            mov input.ki.wScan,ax
            mov input.ki.dwFlags,KEYEVENTF_SCANCODE + KEYEVENTF_KEYUP;;KEYEVENTF_UNICODE   ;KEYEVENTF_SCANCODE + KEYEVENTF_KEYUP ;
            .if bExtendedkey
                mov input.ki.dwFlags,KEYEVENTF_SCANCODE + KEYEVENTF_KEYUP + KEYEVENTF_EXTENDEDKEY
            .endif
            mov input.ki.time,0
            Invoke  GetMessageExtraInfo
            mov input.ki.dwExtraInfo,eax
            mov input.ddType,INPUT_KEYBOARD
            Invoke SendInput,1,addr input,sizeof INPUT   ;keybd_event
            Invoke Sleep,nSleep
        .endif
    ret
SendKey                         endp
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
; #########################################################################
                    ;Fin
; #########################################################################
; «««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
end start
