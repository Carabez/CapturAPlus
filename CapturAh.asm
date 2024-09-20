; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
;                     COMPILER DIRECTIVES
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
.586
.model flat,stdcall
option casemap:none ; case sensitive
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
;                     INCLUDE FILES
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
    Include                     \masm32\include\windows.inc
    Include                     \masm32\include\masm32.inc
    ;
    Include                     \masm32\include\kernel32.inc
    Include                     \masm32\include\user32.inc
    include                     \masm32\include\winmm.inc
    include                     \masm32\include\advapi32.inc
    ;include                     \masm32\include\masm32rt.inc
    ;
    include                     \masm32\macros\macros.asm

    ;
    includelib                  \masm32\lib\kernel32.lib
    includelib                  \masm32\lib\user32.lib
    includelib                  \masm32\lib\winmm.lib
    includelib                  \masm32\lib\advapi32.lib
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;GLOBAL FUNCTIONS
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
;
    DllMain                     PROTO :DWORD,:DWORD,:DWORD
    InstallHook                 PROTO :DWORD,:DWORD,:DWORD
    MouseProc                   PROTO :DWORD,:DWORD,:DWORD
    UninstallHook               PROTO
    a2dw                        PROTO :DWORD
.const
    WM_MOUSEHOOK                equ WM_USER+492
align 8
.data
    hInstance                   dd 0   
.data?
    hEventHook                  dd ?
    hCallBackWin                dd ?
    dEventHookID                dd ?
    dMouseClickDelayTime        dd ? 
    dMenuShowDelay              dd ?    
    ;
    hWinClicking                dd ?
    dLastClickedTime            dd ?
    ;
    dDoubleClickHeight          dd ?
    dDoubleClickWidth           dd ?
    dDoubleClickSpeed           dd ?
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;EntryPoint  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
.code
;DllEntry                        proc hInst:HINSTANCE, reason:DWORD, reserved1:DWORD
align 8
DllMain                         proc hInstDLL:HINSTANCE,reason:DWORD,reserved1:DWORD
                                LOCAL hReg:DWORD
                                LOCAL dSize:DWORD
                                LOCAL szTexto[MAX_PATH]:BYTE
                                ;
    .if reason==DLL_PROCESS_ATTACH
        mov eax,hInstDLL
        mov hInstance,eax
        Invoke DisableThreadLibraryCalls,hInstDLL
        ;
        szText szMouseKey,"Control Panel\Mouse"
        szText szDoubleClickWidth,"DoubleClickWidth"
        szText szDoubleClickHeight,"DoubleClickHeight"
        szText szDoubleClickSpeed,"DoubleClickSpeed"
        ;szText szIntFormat,"%lu"  ;To convert dw to ascci
        Invoke RegOpenKeyEx,HKEY_CURRENT_USER,addr szMouseKey,0,KEY_READ,addr hReg ;HKEY_CURRENT_USER HKEY_LOCAL_MACHINE KEY_ALL_ACCESS
        .if eax!=ERROR_SUCCESS
            mov dDoubleClickHeight,4
            mov dDoubleClickWidth,4
            mov dDoubleClickSpeed,500 ;According to Microsoft's MSDN website, the default timing in Windows is 500 ms
        .else
            mov dSize,sizeof szTexto
            Invoke RegQueryValueEx,hReg,addr szDoubleClickSpeed,0,NULL,addr szTexto,addr dSize
            ;Invoke MessageBox,NULL,addr szDoubleClickSpeed,addr szTexto,MB_OK
            Invoke a2dw,addr szTexto
            cmp eax,50
            jl @F
                mov eax,500
            @@:
            cmp eax,2000
            jg @F
                mov eax,500
            @@:
            mov dDoubleClickSpeed,eax
            ;Invoke wsprintf,addr szTexto,addr szIntFormat,dDoubleClickSpeed
            ;Invoke MessageBox,NULL,addr szDoubleClickSpeed,addr szTexto,MB_OK
            Invoke RegQueryValueEx,hReg,addr szDoubleClickWidth,0,NULL,addr szTexto,addr dSize
            ;Invoke MessageBox,NULL,addr szDoubleClickSpeed,addr szTexto,MB_OK
            Invoke a2dw,addr szTexto
            cmp eax,4
            jl @F
                mov eax,4
            @@:
            cmp eax,32
            jg @F
                mov eax,4
            @@:
            mov dDoubleClickWidth,eax
            ;Invoke MessageBox,NULL,addr szDoubleClickWidth,addr szTexto,MB_OK
            Invoke RegQueryValueEx,hReg,addr szDoubleClickHeight,0,NULL,addr szTexto,addr dSize
            ;Invoke MessageBox,NULL,addr szDoubleClickSpeed,addr szTexto,MB_OK
            Invoke a2dw,addr szTexto
            cmp eax,4
            jl @F
                mov eax,4
            @@:
            cmp eax,32
            jg @F
                mov eax,4
            @@:
            mov dDoubleClickHeight,eax
            ;Invoke MessageBox,NULL,addr szDoubleClickHeight,addr szTexto,MB_OK
            Invoke RegCloseKey,hReg
        .endif
        szText szDesktopKey,"Control Panel\Desktop"
        szText szMenuShowDelay,"MenuShowDelay"        
        Invoke RegOpenKeyEx,HKEY_CURRENT_USER,addr szDesktopKey,0,KEY_READ,addr hReg ;HKEY_CURRENT_USER HKEY_LOCAL_MACHINE KEY_ALL_ACCESS
        .if eax!=ERROR_SUCCESS
            mov dDoubleClickHeight,4
            mov dDoubleClickWidth,4
            mov dDoubleClickSpeed,500 ;According to Microsoft's MSDN website, the default timing in Windows is 500 ms
        .else
            mov dSize,sizeof szTexto
            Invoke RegQueryValueEx,hReg,addr szMenuShowDelay,0,NULL,addr szTexto,addr dSize
            ;Invoke MessageBox,NULL,addr szMenuShowDelay,addr szTexto,MB_OK
            Invoke a2dw,addr szTexto
            cmp eax,1
            jl @F
                mov eax,400 ;Type in a number between 0 to 4000 (400 is default)
            @@:
            cmp eax,4000
            jg @F
                mov eax,400
            @@:
            mov dMenuShowDelay,eax        
        .endif    
        Invoke timeGetTime
        mov dLastClickedTime,eax
    ;.elseif reason==DLL_PROCESS_DETACH
    ;.elseif reason==DLL_THREAD_ATTACH
    ;.else        ; DLL_THREAD_DETACH
    .endif
    mov  eax,TRUE
    ret
DllMain                         endp
;DllEntry Endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;MouseProc  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 8
MouseProc proc                  nCode:DWORD,wParam:DWORD,lParam:DWORD
                                ;LOCAL szTexto[MAX_PATH]:BYTE
                                ;LOCAL dCurrentClicTime:DWORD ;dLastClickedTime
                                ;LOCAL hWhoIsClicking:DWORD ;hWinClicking

;Only windows that have the CS_DBLCLKS style can receive WM_LBUTTONDBLCLK messages,
;which the system generates whenever the user presses, releases, and again presses the left mouse button
;within the system's double-click time limit.
;Double-clicking the left mouse button actually generates a sequence of four messages:
;WM_LBUTTONDOWN, WM_LBUTTONUP, WM_LBUTTONDBLCLK, and WM_LBUTTONUP.
;
;
;Maximum height above/below which 2 clicks will not be considered as a double-click
;HKEY_CURRENT_USER\Control Panel\Mouse\DoubleClickHeight ;
;Delay in milliseconds the system must wait between 2 clicks fro them to be considered a double-click.
;HKEY_CURRENT_USER\Control Panel\Mouse\DoubleClickSpeed
;Maximum Width left/rigth which 2 clicks will not be considered as a double-click
;HKEY_CURRENT_USER\Control Panel\Mouse\DoubleClickWidth
    ;=============================================================
    ;LLAMAR ANTES
    .if nCode != HC_ACTION
        ;If code is less than zero, the hook procedure must pass the message to the CallNextHookEx function
        ; without further processing and should return the value returned by CallNextHookEx.
        Invoke CallNextHookEx,hEventHook,nCode,wParam,lParam
        ;
    .else
        ;If code is HC_ACTION, the hook procedure must process the message.    
        .if (wParam == WM_LBUTTONDOWN || wParam == WM_NCLBUTTONDOWN) ;|| \
            ;(wParam == WM_RBUTTONUP || wParam == WM_NCRBUTTONUP)
            Invoke timeGetTime
            mov ecx,dLastClickedTime
            mov dLastClickedTime,eax
            sub eax,ecx
            .if eax > dDoubleClickSpeed
                    mov edx,lParam
                    assume edx:PTR MOUSEHOOKSTRUCT
                    Invoke WindowFromPoint,[edx].pt.x,[edx].pt.y
                    assume edx:nothing
                    mov hWinClicking,eax
                    Invoke PostMessage,hCallBackWin,dEventHookID,wParam,hWinClicking
                    ;Permitir tiempo para que se capture la ventana ANTES de que se envie el CLICK
                    .if dMouseClickDelayTime != NULL
                        Invoke Sleep,dMouseClickDelayTime
                    .endif
            .endif
            Invoke CallNextHookEx,hEventHook,nCode,wParam,lParam
            ;It is highly recommended that you call CallNextHookEx and return the value it returns
        .elseif (wParam == WM_RBUTTONDOWN || wParam == WM_NCRBUTTONDOWN)
            Invoke CallNextHookEx,hEventHook,nCode,wParam,lParam
            push eax
            ;It is highly recommended that you call CallNextHookEx and return the value it returns
            Invoke timeGetTime
            mov ecx,dLastClickedTime
            mov dLastClickedTime,eax
            sub eax,ecx
            .if eax > dDoubleClickSpeed
                    ;
                    ;Permitir tiempo para que se dibuje el Men· Contextual
                    .if dMouseClickDelayTime != NULL
                        Invoke Sleep,dMouseClickDelayTime
                    .endif                    
;                    mov ecx,dMenuShowDelay
;                    .if dMouseClickDelayTime != NULL
;                        add ecx,dMouseClickDelayTime
;                    .else
;                        add ecx,200                                            
;                    .endif                      
;                    Invoke Sleep,ecx                    
                    ;
                    mov edx,lParam
                    assume edx:PTR MOUSEHOOKSTRUCT
                    Invoke WindowFromPoint,[edx].pt.x,[edx].pt.y
                    assume edx:nothing
                    mov hWinClicking,eax
                    Invoke PostMessage,hCallBackWin,dEventHookID,wParam,hWinClicking
            .endif
            pop eax 
       .else
           Invoke CallNextHookEx,hEventHook,nCode,wParam,lParam
           ;It is highly recommended that you call CallNextHookEx and return the value it returns
        .endif
        ;Invoke CallNextHookEx,hEventHook,nCode,wParam,lParam
        ;It is highly recommended that you call CallNextHookEx and return the value it returns
    .endif
    ret
MouseProc                       endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;InstallHook  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 8
InstallHook                     proc hWnd:DWORD,dEventID:DWORD,dMouseClickSnapDelay:DWORD
    mov eax,hWnd
    mov hCallBackWin,eax
    mov ecx,dEventID
    .if ecx==NULL
        mov ecx,WM_MOUSEHOOK
    .endif
    mov dEventHookID,ecx
    mov edx,dMouseClickSnapDelay
    mov dMouseClickDelayTime,edx
    ;WH_CALLWNDPROC Installs a hook procedure that monitors messages before the system sends them to the destination window procedure. For more information, see the CallWndProc hook procedure.
    ;WH_CALLWNDPROCRET Installs a hook procedure that monitors messages after they have been processed by the destination window procedure. For more information, see the [HOOKPROC callback function](nc-winuser-hookproc.md) hook procedure.
    ;WH_CBT The system calls this function before activating, creating, destroying, minimizing, maximizing, moving, or sizing a window; before completing a system command; before removing a mouse or keyboard event from the system message queue; before setting the keyboard focus; or before synchronizing with the system message queue
    ;WH_DEBUG Installs a hook procedure useful for debugging other hook procedures. For more information, see the DebugProc hook procedure.
    ;WH_FOREGROUNDIDLE Installs a hook procedure that will be called when the application's foreground thread is about to become idle. This hook is useful for performing low priority tasks during idle time.
    ;WH_GETMESSAGE Installs a hook procedure that monitors messages posted to a message queue. For more information, see the GetMsgProc hook procedure.
    ;WH_KEYBOARD Installs a hook procedure that monitors keystroke messages. For more information, see the KeyboardProc hook procedure.
    ;WH_KEYBOARD_LL Installs a hook procedure that monitors low-level keyboard input events. For more information, see the [LowLevelKeyboardProc](/windows/win32/winmsg/lowlevelkeyboardproc) hook procedure.
    ;WH_MOUSE Installs a hook procedure that monitors mouse messages. For more information, see the MouseProc hook procedure.
    ;WH_MOUSE_LL Installs a hook procedure that monitors low-level mouse input events. For more information, see the LowLevelMouseProc hook procedure.
    ;WH_MSGFILTER Installs a hook procedure that monitors messages generated as a result of an input event in a dialog box, message box, menu, or scroll bar. For more information, see the MessageProc hook procedure.
    ;WH_SHELL Installs a hook procedure that receives notifications useful to shell applications. For more information, see the ShellProc hook procedure.
    ;WH_SYSMSGFILTER Installs a hook procedure that monitors messages generated as a result of an input event in a dialog box, message box, menu, or scroll bar. The hook procedure monitors these messages for all applications in the same desktop as the calling thread. For more information, see the SysMsgProc hook procedure.    
    Invoke SetWindowsHookEx,WH_MOUSE_LL,addr MouseProc,hInstance,NULL
    ;Invoke SetWindowsHookEx,WH_MOUSE,addr MouseProc,hInstance,NULL
    mov hEventHook,eax
    ret
InstallHook                     endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;UninstallHook  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 8
UninstallHook                   proc
    invoke UnhookWindowsHookEx,hEventHook
    ret
UninstallHook                   endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 8
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
End DllMain
