_Invoke MACRO funcname:REQ,p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12,p13,p14,p15,p16,p17,p18,p19,p20
LOCAL pos,counter
counter=0

    FOR arg,<p20,p19,p18,p17,p16,p15,p14,p13,p12,p11,p10,p9,p8,p7,p6,p5,p4,p3,p2,p1>

        IFNB <arg>

		counter=counter+4
		pos=@InStr(1,arg,<ADDR >) OR @InStr(1,arg,<addr >) OR @InStr(1,arg,<Addr >)  OR \
                @InStr(1,arg,<ADDR	>) OR @InStr(1,arg,<addr	>) OR @InStr(1,arg,<Addr	>)

             IF pos

                IF ((OPATTR(@SubStr(arg,%pos+5))) EQ 98) OR ((OPATTR(@SubStr(arg,%pos+5))) EQ 34)
				
                        lea eax,@SubStr(<arg>,%pos+5)
                        push eax
                ELSE
                        push OFFSET @SubStr(<arg>,%pos+5)
                ENDIF

             ELSE
                        push arg
             ENDIF
        ENDIF
    ENDM
call funcname

IF ((OPATTR(funcname)) AND 11100000000y) EQ 00100000000y	; Handle C functions
    IF counter NE 0
        add esp,counter						; correct the stack
    ENDIF
ENDIF

ENDM