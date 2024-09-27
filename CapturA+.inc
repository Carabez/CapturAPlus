; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
;                     INCLUDE FILES
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
Include     .\_invoke.inc
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;GLOBAL FUNCTIONS
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
;=================================================
;                       CapturA+.Dll
;=================================================
CapturA_LoadLibrary             PROTO
CapturA_FreeLibrary             PROTO

InstallHook                     PROTO :DWORD,:DWORD,:DWORD
UninstallHook                   PROTO

.const
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;EQUATES
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
;
AA_ERROR_DLL_LOADING_ERROR      equ -001    ;Dll loading error.
AA_ERROR_DLL_PROC_ERROR         equ -002    ;Dll Internal Procedure no found.
AA_ERROR_NULL_PARAM             equ -003    ;Null parameter is NO allowed.
AA_ERROR_INVALID_PARAM          equ -004    ;One or more parameters are not valid.
AA_ERROR_INVALID_BUFFER         equ -005    ;A buffer is not valid.
AA_ERROR_NOT_IMPLEMENTED        equ -006    ;Called Function or parameter was NO implemented.      
;
.data
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;INITIALIZED DATA
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
;=================================================
;                       Captura.Dll
;=================================================
align 4
;szCapturaDll                    db "captura+"               ,0
;
szCapturaDll_InstallHook        db "InstallHook"            ,0
szCapturaDll_UninstallHook      db "UninstallHook"          ,0
;
align 4
hCapturaDll                     dd 0
_CapturaDll                     STRUCT 4    ;align 4
                                    InstallHook                    dd 0
                                    UninstallHook                  dd 0
_CapturaDll                     ENDS
CapturaDll                      _CapturaDll <>
.data?
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;UNINITIALIZED DATA
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
szCapturaDll                    db MAX_PATH dup(?) ;Extracts "CapturA+" from Path.
.code
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
; #########################################################################
                    ;CapturA+.Dll Recalling Functions
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
CapturA_LoadLibrary             proc
        ;
        ;If the string specifies a module name without a path and the file name extension is omitted, the function appends the default library extension ".DLL" to the module name.
        ;To prevent the function from appending ".DLL" to the module name, include a trailing point character (.) in the module name string.
        Invoke LoadLibrary,addr szCapturaDll
        mov hCapturaDll,eax                   ;If the function fails, the return value is NULL; If the function succeeds, the return value is a handle to the module.
    ret
CapturA_LoadLibrary             endp
align 4
CapturA_FreeLibrary             proc
        ;
        ;The system maintains a per-process reference count on all loaded modules.
        ;Calling LoadLibrary increments the reference count.
        ;Calling the FreeLibrary or FreeLibraryAndExitThread function decrements the reference count.
        ;The system unloads a module when its reference count reaches zero or when the process terminates (regardless of the reference count).
        .if hCapturaDll != NULL
            Invoke FreeLibrary,hCapturaDll
            push eax
            .if eax != NULL
                mov hCapturaDll,NULL
                push edi
                ;
                mov edi, offset CapturaDll  ; address of buffer
                mov ecx, sizeof CapturaDll  ; buffer length
                mov  al, 0          ; load al with character
                rep stosb           ; write character to buffer until ecx = 0
                ;
                pop edi
            .endif
            pop eax
        .endif
        ;
    ret
CapturA_FreeLibrary             endp
align 4
InstallHook                     proc a1:DWORD,a2:DWORD,a3:DWORD
        mov eax,hCapturaDll
        .if eax == NULL
            Invoke CapturA_LoadLibrary
        .endif
        .if eax == NULL
            mov eax,AA_ERROR_DLL_LOADING_ERROR
        .else
            mov eax,CapturaDll.InstallHook
            .if eax == NULL
                Invoke GetProcAddress,hCapturaDll,addr szCapturaDll_InstallHook
                .if eax!=NULL
                    mov CapturaDll.InstallHook,eax
                .endif
            .endif
            .if eax==NULL
                mov eax,AA_ERROR_DLL_PROC_ERROR
            .else
                _Invoke CapturaDll.InstallHook,a1,a2,a3
            .endif
        .endif
    ret
InstallHook                     endp
align 4
UninstallHook                   proc
        ;
        mov eax,hCapturaDll
        .if eax == NULL
            Invoke CapturA_LoadLibrary
        .endif
        .if eax == NULL
            mov eax,AA_ERROR_DLL_LOADING_ERROR
        .else
            mov eax,CapturaDll.UninstallHook
            .if eax == NULL
                Invoke GetProcAddress,hCapturaDll,addr szCapturaDll_UninstallHook
                .if eax!=NULL
                    mov CapturaDll.UninstallHook,eax
                .endif
            .endif
            .if eax==NULL
                mov eax,AA_ERROR_DLL_PROC_ERROR
            .else
                _Invoke CapturaDll.UninstallHook
            .endif
        .endif
    ret
UninstallHook                   endp