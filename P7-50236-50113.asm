; multi-segment executable file template.

data segment 
  
  cells db 127,27,12,123,128,0,2,12,123,128,128,128,128,128,128,128
  ;Strings************************
  menu db "MENU"
  idH db "A    B    C    D"  
  result db "Result"
  formula db "Formula"
  blankString db "      $"

  ;Strings About - (numero de caracteres, x inicial, y inicial, string)
  about1 dw 31,4,2,"This program allows the user to"
  about2 dw 34,3,3,"manipulate cells in a spreadsheet."
  about3 dw 33,4,4,"To do so, in edit mode, place and"
  about4 dw 35,3,5,"press the mouse above a cell to add"
  about5 dw 29,5,6,"a value between -127 and 127."
  about6 dw 37,1,7,"To calculate an operation press above"
  about7 dw 28,6,8,"the formula box and write an"
  about8 dw 23,8,9,"expression of the type:"
  about9 dw 30,5,10,"<cell><operation signal><cell>"
  rafael dw 25,5,13,10h,"Rafael Santos, n",0F8h,"50113"
  diogo dw 25,5,16,10h,"Diogo Simoes,  n",0F8h,"50236"
  
  ;Strings menu - (numero de caracteres, x inicial, y inicial, string)
  top dw 30,5,2,"MicrosSoft Office Excel - 1986"
  import dw 18,11,5,"Import spreadsheet"
  show dw 16,12,8,"Show spreadsheet"
  edit dw 16,12,11,"Edit spreadsheet"
  export dw 18,11,14,"Export spreadsheet"
  about dw 5,18,17,"About"
  exit dw 4,18,20,"Exit"
  
  ;Vetores de dimensoes de rectangulos - (canto sup.esq - x, canto sup.esq - y, canto inf.dir - y, canto inf.dir - x)
  tableSize dw 30,20,108,190 ;x0,y0,y1,x1
  formulaSize dw 30,138,158,220
  resultSize dw 240,138,158,310
  menuSize dw 30,170,190,75
  showLines dw 220,20,40,280
  
  importSize dw 40,32, 56, 280 
  showSize   dw 40,56, 80,280
  editSize   dw 40,80,104,280
  exportSize dw 40,104,128,280
  aboutSize  dw 40,128,152,280
  exitSize   dw 40,152,176,280
  
  ;Vetores de dimensoes de linhas verticais - (y inicial, x final, x)
  vl1 dw 20,108,70 ;y0,y1,x
  vl2 dw 20,108,110
  vl3 dw 20,108,150
  ;Vetores de dimensoes de linhas horizontais - (x inicial, x final, y)
  hl1 dw 30,190,42 ;x0,x1,y
  hl2 dw 30,190,64
  hl3 dw 30,190,86
  
  ;Outras Variaveis
  activePage db 0
  ascii db ' ',' ',' ',' '
  ascii16 db ' ',' ',' ',' ',' ',' ','$'
  
  AskOpen dw "Opening file name:"
  Ask dw "Enter the name of the text file:"           
  Contents db "Contents.bin", 0                       ;Nome do ficheiro Contents.bin
  FileHandle dw 0                                     ;FileHandle do ficheiro a criar
  Filename db 22 dup (?)                              ;String com o nome do ficheiro a inserir pelo utilizador
  TemporaryBuffer db 137 dup (?)
  formulaBuffer db 9 dup (?) 
  formulaReady db 5 dup (' '), '$'
         
ends

stack segment
    dw   128  dup(0)
ends

code segment
start:
; set segment registers:
    mov ax, data
    mov ds, ax
    mov es, ax 
    
    call loadContents
    
    menuLabel:
    mov activePage, 0
    mov al, 0h
    call setVideoMode
    call writeMenu
    jmp repeatVerification
    
    showLabel:
    mov activePage, 2
    jmp drawScreenLabel
    
    editLabel:
    mov activePage, 3
    drawScreenLabel:                 
    mov al, 13h                 
    call setVideoMode
    
    ;call drawScreen
    call writebutton
    call writeH
    call writeV  
    call writeResult
    call writeFormula
    call writeCells
    call writeFormulaBox
    call calculator
    jmp repeatVerification
    
    formulaLabel:  
    call wInFormula
    call calculator
    jmp repeatVerification
    
    cellLabel:; registers in use - cx, dx
    call findCell  ; bx (bl)
    mov DI, offset cells
    add DI, bx     ; DI
    push BX                     
    call setCoordinatesCell ; dx (dh,dl)
    call setCursorPosition
    pop BX
    call setValue
    or bl, bl
    jnz repeatVerification
    call convertNumber8bit ; DI
    call printNumber
    call calculator    
    jmp repeatVerification
    
    importLabel:
    call importFile
    jmp menuLabel
    
    exportLabel:
    call exportFile
    jmp menuLabel
    
    aboutLabel:
    call clearScreen
    mov SI, offset about1
    call writeString
    mov SI, offset about2
    call writeString
    mov SI, offset about3
    call writeString
    mov SI, offset about4
    call writeString
    mov SI, offset about5
    call writeString
    mov SI, offset about6
    call writeString
    mov SI, offset about7
    call writeString
    mov SI, offset about8
    call writeString
    mov SI, offset about9
    call writeString 
    mov SI, offset rafael
    call writeString
    mov SI, offset diogo
    call writeString
    call hideBlinkingCursor 
    mov ah, 7 
    
    ver:    
    int 21H                    
    cmp al, 1BH                
    je exitAbout               
    cmp al, 0DH                
    jne ver                  
    
    exitAbout:
    jmp menuLabel
    
    repeatVerification:
    call getMousePos
    mov SI, offset activePage
    call verifyMousePosition
    cmp bl, 0
    je repeatVerification
    cmp bl, 1
    je menuLabel
    cmp bl, 2
    je formulaLabel
    cmp bl, 19
    jb cellLabel
    je importLabel
    cmp bl, 20
    je editLabel
    cmp bl, 21
    je showLabel
    cmp bl, 22
    je exportLabel
    cmp bl, 23
    je aboutLabel
    
    exitLabel:
    
    mov dx, offset Contents  ;Passagem para o dx do offset da string que contem "contents.bin"
    call fcreate             ;Destruicao e criacao do ficheiro contents.bin 
    
    mov al, 1                ;AL com 1 = Modo de abertura de ficheiro : write
    call fopen               ;Abre o ficheiro contents.bin
    
    mov bx, ax               ;Move para o BX, o file handle
    
    mov cx, 16               ;Numero de bytes a escrever
    mov dx, offset cells     ;Offset do primeiro byte a escrever
    call fwrite
    mov dx, offset formulaReady
    call fwrite              
    call fclose              
    
    mov ax, 4C00H            
    int 21H                        
ends
    
;*************************************************************
;Initialize Mouse
;
;Input:Nothing
;Output:
; AX- 0000h if Error; FFFFh if Detected
; BX=number of mouse buttons
;Destroys:Nothing
;************************************************************* 
  
  initMouse proc
    
    mov ax,00
    int 33h
  
  ret
  initMouse endp
  
;*************************************************************
;Show Mouse Pointer
;
;Input: Nothing
;Output: Nothing
;*************************************************************
  
  showMouse proc  
    
    push ax
    mov ax,01
    int 33h
    pop ax
  
    ret
  showMouse endp
  
;*************************************************************
;Hide Mouse Pointer
;
;Input: Nothing
;Output: Nothing
;*************************************************************

  hideMouse proc
  
    push ax
    mov ax,02
    int 33h
    pop ax
  
    ret
  hideMouse endp
  
;*************************************************************
;Get Mouse Position and Button pressed 
;
;Input: Nothing
;Output:
; BX- Button pressed (1 - botao da esquerda, 2 - botao da direita e 3 ambos os botoes)
; CX- horizontal position (column)
; DX- Vertical position (row)
;Destroys: Nothing else
;*************************************************************
 
  getMousePos proc
  
    push ax
    mov ax,03h
    int 33h
    pop ax
  
    ret
  getMousePos endp
  
;*************************************************************
;Set Video Mode
;
;Input:
; AL- Video Mode
; - 00h - text mode. 40x25. 16 colors. 8 pages.
; - 03h - text mode. 80x25. 16 colors. 8 pages.
; - 13h - graphical mode. 40x25. 256 colors. 320x200 pixels, 1 page.
;Output:Nothing
;Detroys: Nothing 
;*************************************************************

  setVideoMode proc
    
    push ax
    mov ah,00 ; modo video
    int 10h
    pop ax
    
    ret
  setVideoMode endp
  
;*************************************************************
;hideBlinkingCursor
;Description:makes cursor invisible
;Input: 
; CH: start of cursor (bits 0-4), options (bits 5-7)
; CL: end of cursor (bits 0-4), Bit 5 de CH a 1 hides cursor
;Output:Nothing
;Destroys:Nothing
;*************************************************************  
  
  HideBlinkingCursor proc 
    
    push cx
    mov ch, 32 ;(0001 0000B)
    mov ah, 1
    int 10h
    pop cx
  
    ret
  HideBlinkingCursor endp 

;*************************************************************
;showStandardCursor
;
;Description:shows a standard cursor in text mode
;Input:               
; CH: start of cursor (bits 0-4), options (bits 5-7)
; CL: end of cursor (bits 0-4), Bit 5 de CH a 1 hides cursor
;Output:Nothing
;Destroys:Nothing
;*************************************************************
  
  ShowStandardCursor proc
    push cx
    mov ch, 6
    mov cl, 7
    mov ah, 1
    int 10h
    pop cx
    
    ret
  ShowStandardCursor endp
  
;*************************************************************
;Set Cursor Position
;
;Input:
; DH = row.
; DL = column.
; BH = page number (0..7).
;Output:Nothing
;Detroys: Nothing
;*************************************************************
  setCursorPosition proc
  
    mov ah,2
    int 10h
    
    ret
  setCursorPosition endp
  
;*************************************************************
;Set Cursor Position + Size
;
;Input:
; BH = page number (0..7).
;Output:
; DH = row. DL = column. CH = Cursor start line. CL = Cursorbottom line.
;Detroys: Nothing
;*************************************************************
  
  GetCursorPosition proc
  
    mov ah,3
    int 10h
  
    ret
  GetCursorPosition endp
  
;*************************************************************
;Select active page
;
;Input:
; AL = page number (0..7).
;Output: Nothing
;Detroys: Nothing
;*************************************************************
  
  SelectActivePage proc
  
    mov ah,5
    int 10h
  
    ret
  SelectActivePage endp
  
; ************************************************************
;Clear Screen
;
;Input:Nothing
;Output:Nothing
;Detroys:Nothing
;*************************************************************
  
  clearScreen proc
    push ax
    mov ah,06
    mov al,00
    mov BH,07       ; attributes to be used on blanked lines
    mov cx,0        ; CH,CL = row,column of upper left corner of window to scroll
    mov DH,25       ;= row,column of lower right corner of window
    mov DL,40
    int 10h
    pop ax
  
    ret
  clearScreen endp
  
;************************************************************
;Put Pixel
;
;Input:
; AL- pixel value
; CX- column
; DX- row
;Output:Nothing
;Detroys: Nothing
;************************************************************
  
  putPixel proc
    push ax
    push bx
    mov ah,0ch
    mov bh,00 ; active display page
    int 10h
    pop bx
    pop ax
  
    ret
  putPixel endp
  
;************************************************************
;Get Video Mode
;
;Input:Nothing
;Output:
; AL- Video Mode
; BH- Display Page
;Destroys: AH
;************************************************************ 
  
  getVideoMode proc
    
    mov ah,0fh ; get video mode
    int 10h
  
    ret
  getVideoMode endp
  
;************************************************************
;Draw Screen
; 
;Description: calls the functions to write the video mode design
;Input:Nothing
;Output: Nothing
;Destroys: SI
;************************************************************   
    
  drawScreen proc

    mov al, 0Fh  
    
    mov SI, offset tableSize
    call drawRect
    mov SI, offset vl1
    call drawVLine
    mov SI, offset vl2
    call drawVLine
    mov SI, offset vl3
    call drawVLine
    mov SI, offset hl1
    call drawHLine
    mov SI, offset hl2
    call drawHLine
    mov SI, offset hl3
    call drawHLine
    
    mov SI, offset formulaSize
    call drawRect
    
    mov SI, offset resultSize
    call drawRect
    
    mov SI, offset menuSize
    call drawRect
    
    ret 
  drawScreen endp  
  
  
;************************************************************
;Draw Rect
; 
;Description: draws a rectangle in video mode
;Input:
; SI - offset coordenadas limitativas de um rectangulo
;Output:Nothing
;Destroys: CX, DX
;************************************************************
  
  drawRect proc
    
    mov cx, [SI] ;x0
    mov dx, [SI+2] ;y0
    
    linhaEsquerda: 
    call putPixel
    inc dx
    cmp dx, [SI+4] ;y1
    jne linhaEsquerda
     
    linhaInferior:
    call putPixel
    inc cx
    cmp cx, [SI+6] ;x1
    jne linhaInferior 
    
    linhaDireita:
    call putPixel
    dec dx
    cmp dx, [SI+2] ;y0
    jne linhaDireita
    
    linhaSuperior:
    call putPixel
    dec cx
    cmp cx, [SI] ;x0
    jne linhaSuperior 
    
    ret
  drawRect endp  
  
;************************************************************
;drawHLine
;    
;Description: draws horizontal lines in video mode
;Input:
; SI - offset coordenadas linhas horizontais
;Output:Nothing
;Destroys: CX, DX
;************************************************************ 
  
  drawHLine proc
    
    mov cx, [SI]      ;x0
    mov dx, [SI+4]    ;y 
    
    hLine: 
    call putPixel
    inc cx
    cmp cx, [SI+2]    ;x1
    jne hLine
    
    ret
  drawHLine endp
  
;************************************************************
;drawVLine
;         
;Description: draws vertical lines in video mode
;Input:
; SI - offset coordenadas linhas verticais
;Output:Nothing
;Destroys: CX, DX
;************************************************************

  drawVLine proc
    
    mov cx, [SI+4]    ;x
    mov dx, [SI]      ;y0
    
    vLine: 
    call putPixel
    inc dx
    cmp dx, [SI+2]    ;y1
    jne vLine
    
    ret      
  drawVLine endp
  
;************************************************************
;writeButton
;            
;Description: writes menu in video mode
;Input: 
; menu = string that contains the word menu
;Output:Nothing
;Destroys: AX, BX, CX, DX
;************************************************************
  
  writeButton proc
    
    mov ax, 1301h
    mov bh, 0         ;page
    mov bl, 0Fh       ;atribute
    mov cx, 4         ;number of chars
    mov dh, 22        ;starting y
    mov dl, 5         ;starting x
    push ds
    pop es
    mov bp, offset menu
    int 10h
    
    ret             
  writeButton endp 
  
;************************************************************
;writeFormula
;
;Description: writes formula in video mode
;Input: 
; formula = string that contains the word menu
;Output:Nothing
;Destroys: AX, BX, CX, DX
;************************************************************
  
  writeFormula proc          ;deveria ser otimizado
                   
    mov ax, 1301h
    mov bh, 0
    mov bl, 0Fh
    mov cx, 7
    mov dh, 15
    mov dl, 5
    push ds
    pop es
    mov bp, offset formula
    int 10h
    
    ret                                      
  writeFormula endp
  
;************************************************************
;writeResult
;
;Description: write Result in video mode
;Input: 
; result = string that contains the word menu
;Output:Nothing
;Destroys: AX, BX, CX, DX
;************************************************************
  
  writeResult proc      ;deveria ser otimizado
    
    mov ax, 1301h
    mov bh, 0
    mov bl, 0Fh
    mov cx, 6
    mov dh, 15
    mov dl, 30
    push ds
    pop es
    mov bp, offset result                                           
    int 10h
    
    ret
  writeResult endp

;************************************************************
;writeButton
;
;Input: 
; idH = string that contains the line to numers of the collumns
;Output:Nothing
;Destroys: AX, BX, CX, DX
;************************************************************
  
  writeH proc                                      
    
    mov ax, 1301h
    mov bh, 0
    mov bl, 0Fh
    mov cx, 16
    mov dh, 1
    mov dl, 6
    push ds
    pop es
    mov bp, offset idH
    int 10h
    
    ret
  writeH endp

;************************************************************
;writeV 
;
;Description: draws an horizontal line
;Input:Nothing
;Output:Nothing
;Destroys: AX, BX, CX, DX
;************************************************************
  
  writeV proc 
    
    mov dl, 2   ;starting x 
    mov bh, 0   ;page number
    mov dh, 3   ;starting y   
    mov al, 31h ;character
    mov bl, 0Fh ;atribute 
    mov cx, 1   ;number of times to write the char
    
    character:         
    mov ah, 2 
    int 10h                    
    mov ah, 09h            
    add dh, 3 
    int 10h
    inc al
    cmp al, 35h
    jne character
                   
    ret
  writeV endp 
  
;************************************************************
;writeCells 
;
;Description: writes a number in a cell
;Input:DI - number of the cell
;Output:Nothing
;Destroys: BX
;************************************************************
   
  writeCells proc
    
    push 0
    mov bx, 0
    
    repeatCell:
    mov DI, bx     ; DI
    cmp [DI], 128
    je jumpPrint
    call setCoordinatesCell ; dx (dh,dl)
    call convertNumber8bit ; DI
    call printNumber
    
    jumpPrint:
    pop bx
    inc bx
    push bx
    cmp bx, 16
    jne repeatCell
    pop bx
    
    ret
  writeCells endp
    
;***********************************************************
;verifyMousePosition
;
;verifies if the mouse is pressing inside any cell,
;calling funtions if it is pressed at all, functions for each cell
;
;Input: SI-offset of activePage
;       bl-mouse pressed
;       CX-horizontal position
;       DX-vertical position
;Output: bl-0-nothing was clicked
;          -1-menu
;          -2-formula
;          -3-cell  A1(0)     -11-cell C1(8)
;          -4-cell  A2(1)     -12-cell C2(9)
;          -5-cell  A3(2)     -13-cell C3(10)
;          -6-cell  A4(3)     -14-cell C4(11)
;          -7-cell  B1(4)     -15-cell D1(12)
;          -8-cell  B2(5)     -16-cell D2(13)
;          -9-cell  B3(6)     -17-cell D3(14)
;          -10-cell B4(7)     -18-cell D4(15)               
;          -19-import
;          -20-show
;          -21-edit
;          -22-export
;          -23-about
;          -24-exit 
;Destroys: BX, CX, SI  
;**************************** a lista esta ordenada por ordem de utilizacao do mais
;**************************** utilizado (0) para o menos utilizado (24), aproximadamente
  
  verifyMousePosition proc
    
    mov bh, [SI]
    cmp bl, 1
    je jump
    mov bl, 0
    jmp endOfVerification
    
    jump: 
    shr cx, 1 
     
    cmp bh, 0
    je page0
    
    mov SI, offset menuSize        ;SHOW & EDIT    
    call verifyBox
    cmp bl, 1
    je endOfVerification
    
    cmp bh, 2
    je page2
    jmp endOfVerification
    
    page0:                         ;MENU
    mov SI, offset  importSize
    call verifyBox
    cmp bl, 1
    jne notImport
    mov bl, 19
    jmp endOfVerification  
    notImport:
    
    mov SI, offset  showSize
    call verifyBox    
    cmp bl, 1
    jne notShow
    mov bl, 20
    jmp endOfVerification  
    notShow:
    
    mov SI, offset  editSize                
    call verifyBox  
    cmp bl, 1
    jne notEdit
    mov bl, 21
    jmp endOfVerification  
    notEdit:
    
    mov SI, offset  exportSize
    call verifyBox   
    cmp bl, 1
    jne notExport
    mov bl, 22
    jmp endOfVerification  
    notExport:
    
    mov SI, offset  aboutSize
    call verifyBox       
    cmp bl, 1
    jne notAbout
    mov bl, 23
    jmp endOfVerification  
    notAbout:
    
    mov SI, offset  exitSize
    call verifyBox
    cmp bl, 1
    jne notExit
    mov bl, 24  
    notExit:
    
    jmp endOfVerification
    
     page2:                       ;EDIT
    mov SI, offset formulaSize
    call verifyBox
    cmp bl, 1
    jne notFormula
    mov bl, 2
    jmp endOfVerification
    notFormula:
    
    mov SI, offset tableSize
    call verifyBox
    cmp bl, 1
    jne endOfVerification
    mov bl, 3
    
    endOfVerification:
    
    ret
  verifyMousePosition endp
    
;*************************************************************
;verifyBox
;
;Description: verifies if the mouse is within the box with coordenates pointed by SI
;Input: 
; SI-offset of the box coordenates
; CX-mouse horizontal position
; DX-mouse vertical position
;Output: BL-
; -0-the mouse is not inside the box
; -1-the mouse is inside the box             
;Destroys: CX, BL
;*************************************************************
  
  verifyBox proc
    
    mov bl, 0
    
    cmp cx, [SI]
    jb endVerBox
    cmp cx, [SI+6]
    ja endVerBox
    cmp dx, [SI+2]
    jb endVerBox
    cmp dx, [SI+4]
    ja endVerBox
    
    mov bl, 1
    
    endVerBox:
    
    ret
  verifyBox endp 
 
;*************************************************************************
;printfBlank
;
;Description: prints on the screen four blank characters in a line
;Input:
; dh - x of cursor
; dl - y of cursor
;Output: Nothing
;Destroys: CL, SI
;*************************************************************************
  
  printBlank proc
    
    push DX
    
    xor cl, cl             ;
    mov SI, offset ascii   ;
    repeatBlank:           ;
    mov [SI], 20h          ; fills ascii with space characters 
    inc SI                 ;
    inc cl                 ;
    cmp cl, 4              ;
    jne repeatBlank        ;
    
    call printNumber
    
    pop DX
    ret
  printBlank endp    
  
;*************************************************************************
;setValue
;   
;Description: sets the value of a cell to a read number
;Input:
; DI - number of the cell
;output:
; BL - 0 if sucessful
;    - 1 if not sucessful 
;Destroys: BX, CX 
;***************************************************
  
  setValue proc
    
    push DX
    push BX
    
    call printBlank
    call setCursorPosition
    call readNumber
    or bl, bl
    pop BX
    jz errorNotFound
    cmp [DI], 128
    je blankCell
    call setCoordinatesCell ;<----BX
    call convertNumber8bit
    call printNumber
    mov bl, 1
    jmp endSetValue
    
    blankCell:
    call setCoordinatesCell
    call printBlank
    jmp endSetValue
    
    errorNotFound:            
    mov [DI], dl
    cmp [di], 32
    je blankCell
    xor bl, bl
    endSetValue:
    pop DX
    ret
  setValue endp  
  
;*************************************************************
;readNumber
;
;Input:Nothing
;output: 
; bl - 0 if sucessful
;    - 1 if not sucessful
; DL - the read number
;Destroys: AX, BX, CX, DX, SI
;************************************************************* 
  
  readNumber proc
    
    mov SI, offset ascii  ;prepares SI
    xor bl, bl            ;prepares output (bl)
    xor CX, CX            ;prepares counter
    
    repeatReadNumber:
    cmp cl, 4
    jne notLast
    mov ah, 7             ;prepares no echo
    jmp jumpMovAh
    notLast:
    mov ah, 1
    jumpMovAh:            
    int 21h              
    xor ah, ah            ;prepares to push
    cmp al, 0Dh           ;verifies if ENTER
    je endReadNumb
    or cl, cl
    jnz notFirstDigit
    cmp al, 2Dh           ;checks for '-'
    je signaled
    notFirstDigit:
    cmp al, 30h   ;verifies if it is between 0 and 9
    jb notSucessful
    cmp al, 39h
    ja notSucessful
    cmp cl, 4
    je notSucessful
    sub al, 30h ;converts to hexadecimal
    signaled:
    mov [SI], al
    inc SI
    inc cl             
    jmp repeatReadNumber:
    
    notSucessful:
    mov bl, 1
    jmp endConvertion
    
    endReadNumb:
    xor DX, DX
    mov bh, 10
    mov SI, offset ascii
    mov al, [SI]
    inc SI
    cmp al, 2Dh
    jne notSignaled    ; negative
    mov ch, 1
    dec cl
    repeatConvertion:
    mov al, [SI]
    cmp al, 20h
    je endConvertion
    inc SI
    notSignaled:
    add dl, al
    jc notSucessful                           ;NOTA: no import temos uma converção de ascii para unsigned
    cmp [SI], 20h                             ;semelhante a esta, optamos por manter ambos os metodos em vez
    je negation                               ;de apenas este pois o outro metodo é mais simples
    mov al, dl                                ;tendo a desvantagem de não suportar a leitura de sinal
    mul bh                                    ;evitando instruções desnecessarias
    jc notSucessful                           
    mov dl, al
    dec cl
    cmp cl, 1
    ja repeatConvertion
    mov al, [SI]
    add dl, al
    negation:
    cmp dl, 128
    ja notSucessful
    or ch, ch
    jz endConvertion
    neg dl
    endConvertion:
    ret
  readNumber endp
  
;*************************************************************
;findCell  
;
;Description: finds out the number of a cell
;Input:   CX-mouse horizontal position
;         DX-mouse vertical position
;Output:bl-0-A1     4-B1     8-C1     12-D1
;          1-A2     5-B2     9-C2     13-D2
;          2-A3     6-B3     10-C3    14-D3
;          3-A4     7-B4     11-C4    15-D4 
;Destroys: BX
;*************************************************************
  
  findCell proc
    
    cmp cx, 110              
    jb leftx2
    cmp cx, 150
    jb leftx3
    mov bx, 12
    jmp findY
    
    leftx3:
    mov bx, 8
    jmp findY
    
    leftx2:
    cmp cx, 70
    jb leftx1
    mov bx, 4     
    jmp findY
    
    leftx1:
    mov bx, 0
    
    findY:
    cmp dx, 64
    jb upy2
    cmp dx, 86
    jb upy3
    add bl, 3
    jmp found
    
    upy3:
    add bl, 2
    jmp found    
    
    upy2:
    cmp dx, 42
    jb found
    inc bl
                                                                      
    found:
    
    ret
  findCell endp
  
;*********************************************************
;setCoordinatesCell
;
;Description:set the coordinates of cell BL in (dl;dh)
;Input: 
; BX - number of the cell
;Output:
; dh - y
; dl - x
;Destroys: DX, BX, AL      
;*********************************************************
  
  setCoordinatesCell proc
    
    push bx
    and bl, 3
    inc bl
    mov al, bl
    mov bl, 3
    mul bl
    mov dh, al
    pop bx
    and bl, 12
    cmp bl, 0
    jne notA
    mov dl, 4    ;x0
    
    notA:
    cmp bl, 4
    jne notB
    mov dl, 9    ;x1
    
    notB:
    cmp bl, 8
    jne notC
    mov dl, 14   ;x2
    
    notC:
    cmp bl, 12
    jne notD
    mov dl, 19   ;x3
    
    notD:
    
    ret
  setCoordinatesCell endp
  
;************************************************************
;convertNumber8bit  
;
;Description: convert number into ascii form
;Input: 
; DI = number of the cell
;Output:Nothing
;Destroys: AX, BX, CX, DX, DI
;************************************************************
  
  convertNumber8bit proc
    
    push DX         
    xor CX, CX           
    mov BX, 10
    mov al, [DI]
    mov ah, 0 
    
    cmp al, 128
    jb nextDigit
    neg al
    
    nextDigit:
    mov DX, 0
    div BX
    add dl, 30h
    push DX
    inc ch  
    cmp al, 0
    jne nextDigit
    cmp [DI], 128
    jb fill
    push 2Dh
    inc ch
    
    fill:
    mov DX, -1
    repeatFill:
    mov DI, offset ascii
    inc DX
    mov al, 4
    sub al, ch
    add DI, DX
    mov [DI], 20h
    cmp dl, al
    jne repeatFill
    
    positive:
    mov al, 4
    sub al, ch
    mov ah, 0
    mov DI, offset ascii
    add DI, AX
    pop BX               
    mov [DI], bl    
    dec ch
    cmp ch, 0
    jne positive
    pop DX
    
    ret     
  convertNumber8bit endp
  
  
;********************************************************************
;convertNumber16bit
;
;Input:
; AX - number to convert
;Output: Nothing
;Destroys: BX, CX, DI
;********************************************************************  
  
  convertNumber16bit proc
    
    push DX
    push AX         
    xor CX, CX           
    mov BX, 10
    mov cl, 2
    
    cmp AX, 7FFFh
    jb nextDigit16
    neg AX
    
    nextDigit16:
    mov DX, 0
    div BX
    add dl, 30h
    push DX
    inc ch  
    cmp al, 0
    jne nextDigit16 
    mov al, ch
    mul cl
    mov BP, SP
    add BP, AX
    cmp [BP], 0
    jg fill16
    push 2Dh
    inc ch
    
    fill16:
    mov DX, -1
    repeatFill16:
    mov DI, offset ascii16
    inc DX
    mov al, 6
    sub al, ch
    add DI, DX
    mov [DI], 20h
    cmp dl, al
    jne repeatFill16
    
    positive16:
    mov al, 6
    sub al, ch
    mov ah, 0
    mov DI, offset ascii16
    add DI, AX
    pop BX               
    mov [DI], bl    
    dec ch
    cmp ch, 0
    jne positive16
    pop AX
    pop DX
    
    ret
  convertNumber16bit endp  
  
;****************************************************************************
;printNumber
;
;Description: prints the number in the string ascii on the screen
;Input:
; dh - y of cursor
; dl - x of cursor
;Output:Nothing
;Destroys: AX, BX, CX, BP
;****************************************************************************
  
  printNumber proc
    
    push DX
    
    mov AX, 1301h
    mov bh, 0
    mov bl, 0Fh
    mov CX, 4
    push ds
    pop es
    mov BP, offset ascii
    int 10h
    
    pop DX
    ret
  printNumber endp  

;************************************************************
;writeMenu
;
;Description: writes the options in the text menu
;Input:Nothing
;Output:Nothing
;Destroys: SI
;************************************************************
  
  writeMenu proc
    
    call lines
    mov SI, offset top
    call writeString
    mov SI, offset import
    call writeString
    mov SI, offset show
    call writeString
    mov SI, offset edit
    call writeString                            
    mov SI, offset export
    call writeString
    mov SI, offset about
    call writeString
    mov SI, offset exit
    call writeString
    
    ret
  writeMenu endp
  
;************************************************************
;lines
; 
;Description: horizontal lines in text mode
;Input:Nothing
;Output:Nothing
;Destroys: DX, BX, AX, CX
;************************************************************
  
  lines proc
    
    mov dl, 5   ;starting x 
    mov dh, 4   ;starting y
    mov bh, 0   
    mov al, 0C4h
    mov bl, 0Fh  
    mov cx, 30
    
    line:         
    mov ah, 2 
    int 10h                    
    mov ah, 09h            
    add dh, 3 
    int 10h
    cmp dh, 25
    jne line   
    
    ret
  lines endp  

;************************************************************
;writeString
;         
;Description: writes a string in video mode
;Input: 
; SI = string to print and it's coordinates
;Output:Nothing
;Destroys: AX, BX, CX, DX, SI, BP
;************************************************************
  
  writeString proc
    
    mov ax, 1301h
    mov bh, 0
    mov bl, 0Fh
    mov cx, [SI]
    mov dl, [SI+2]
    mov dh, [SI+4]
    push ds
    pop es
    add SI, 6
    mov bp, SI
    int 10h
    
    ret
  writeString endp 

;************************************************************
;checkSignal
;
;Description:
;Input:
;Output:
; bl: 0 - mul
;     1 - add
;     2 - sub
;     3 - div
;     4 - error
;Destroys:
;************************************************************
  checkSignal proc
     
     mov SI, offset formulaBuffer
     add SI, 2
     mov bl, 4
     
     cmp [SI], 'x'
     jne notMultiplication
     mov bl, 0
     jmp notDivision
     notMultiplication:
     
     cmp [SI], '+'
     jne notAddition 
     mov bl, 1
     jmp notDivision
     notAddition:
      
     cmp [SI], '-'
     jne notSubtraction
     mov bl, 2
     jmp notDivision
     notSubtraction:
     
     cmp [SI], '/'
     jne notDivision
     mov bl, 3
     notDivision:
     ret
   checkSignal endp
  
;************************************************************
;wInFormula
; 
;Description: procedure that recongnizes an operation
;Input: 
;Output:Nothing
;Destroys: AX, BX, CX, DX, SI, DI
;************************************************************
  
  wInFormula proc   
     
     mov si, offset formulaBuffer        
     mov dh, 18
     mov dl, 5
     mov bh, 0
     call setCursorPosition      
     
     push dx
     mov cx, 6
     mov dx, offset blankString
     mov ah, 9
     int 21h
     pop dx
     
     call setCursorPosition 
     
     xor cl, cl
                             
     leitura:             
     call getChar
     cmp al, 0Dh
     je check
     cmp cl, 5
     je erro
     call printChar
     inc cl 
     mov [si], al
     inc si
     jmp leitura
     
     check: 
     mov si, offset formulaBuffer 
     
     xor cl, cl
     caracter:                                     
                                                   
     cmp [si], 41h
     jl erro
     
     cmp [si], 44h
     ja erro 
     
     number:
     
     inc si
     cmp [si], 31h
     jl erro
     
     cmp [si], 34h
     ja erro
     
     inc cl
     add si, 2
     
     cmp cl, 1
     je caracter
         
     call checkSignal
     
     cmp bl, 4
     je erro
     
     mov si, offset formulaBuffer
     mov di, offset formulaReady
     mov cx, 5
     rep movsb 
     
     erro:
     
     call setCursorPosition
     mov dx, offset formulaReady
     mov ah, 9
     int 21h
                  
    ret             
  wInFormula endp 
  
;************************************************************
;writeFormulaBox
;
;Input:
; DX - string formulaReady
;Output: Nothing
;Destroys: AH, BH, DX
;************************************************************ 
  
  writeFormulaBox proc
    
    mov dh, 18
    mov dl, 5
    mov bh, 0
    call setCursorPosition
    mov ah, 9
    mov dx, offset FormulaReady
    int 21h
    
    
    ret
  writeFormulaBox endp

;************************************************************
;multiplication
;
;Input:
; AX - 1st cell value
; CX - 2nd cell value
;Output: DXAX - result
;Destroys: AX, DX
;************************************************************ 

  multiplication proc
    
    cmp AX, 0
    jg notNegAX
    cmp CX, 0
    jg notNegAX
    neg AX
    neg CX
    notNegAX:
    imul CX
    ret
  multiplication endp  

;************************************************************
;sum
;
;Input:
; AX - 1st cell value
; CX - 2nd cell value
;Output: AX - result
;Destroys: AX
;************************************************************

  sum proc
    
    add AX, CX
    ret
  sum endp

;************************************************************
;subtraction
;
;Input:
; AX - 1st cell value
; CX - 2nd cell value
;Output: AX - result
;Destroys: AX
;************************************************************  

  subtraction proc
    
    sub AX, CX
    ret
  subtraction endp
  
;************************************************************
;division
;
;Input:
; AX - 1st cell value
; CX - 2nd cell value
;Output:
; DXAX - result
; DX - remainder
;Destroys: AX, CX
;************************************************************  
 
  division proc
    
    xor DX, DX
    cmp al, 0
    jg notNegDX
    mov DX, 0FFFFh
    cmp cl, 0
    jg notNegDX
    neg AX
    neg CX
    xor DX, DX
    notNegDX:
    
    idiv CX
    ret
  division endp
  
;************************************************************
;calculator
;
;Input:
; bl:0-mul
;    1-sum
;    2-sub
;    3-div
;Output:Nothing
;Destroys:
;************************************************************
 
  calculator proc
    
    xor AX, AX
    mov bh, 1
    call getCellFormula    
    mov CX, AX            ;value of second cell in bl
    
    dec bh
    call getCellFormula   ;value of first cell in al
    
    call checkSignal
    
    cmp bl, 4
    je endCalculation
     
    or bl, bl
    jnz notMul
    call multiplication
    jmp endCalculation
    notMul:
    cmp bl, 1
    jne notSum
    call sum
    jmp endCalculation
    notSum:
    cmp bl, 2
    jne notSub
    call subtraction
    jmp endCalculation
    notSub:
    call division    
    endCalculation:
    
    call convertNumber16bit
    mov dl, 32
    mov dh, 18
    mov bh, 0
    call setCursorPosition
    mov DX, offset ascii16
    mov ah, 9
    int 21h
     
    ret
  calculator endp
                                                               
;************************************************************  
;getCellFormula - returns the value in the 1st or 2nd cell of the formula
;
;Input:
; bh - 0 to read the 1st cell
;    - 1 to read the 2nd cell
;Output:
; AX - value in the cell read
;Destroys: AX, SI, DI
;************************************************************
  getCellFormula proc
    
    push DX
    push CX
    push BX
    
    mov SI, offset formulaReady
    xor ah, ah
    mov ch, 3
    mov al, bh      
    mul ch         ;0*3 or 1*3
    add SI, AX     ;1st cell or 2nd cell
    mov al, [SI]   
    cmp al, 41h    
    jne notACell       
    xor al, al     
    notACell:          
    cmp al, 42h    
    jne notBCell       
    mov al, 4      
    notBCell:          
    cmp al, 43h    
    jne notCCell       
    mov al, 8      
    notCCell:
    cmp al, 44h
    jne notDCell          
    mov al, 12
    notDCell:     
    inc SI         
    add al, [SI]
    sub al, 30h     ;add to al(cell value) the number of the line
    dec al          ;dec the cell value because line value starts at 1 instead of 0
    
    mov DI, offset cells
    add DI, AX
    mov al, [DI]
    xor ah, ah
    
    cmp AX, 128
    jb notNegativeValue
    jne not128
    xor AX, AX
    not128:
    CBW
    notNegativeValue:
    
    pop BX
    pop CX
    pop DX
    
    ret   
  getCellFormula endp  
  
;************************************************************
;importFile
;         
;Description: imports a text file that contains the cell and formula information
;Input:Nothing
;Output:Nothing
;Destroys: AX, BX, CX, DX, SI, DI
;************************************************************
  
  importFile proc 
    
    startImport:
    
    mov al, 13H                
    call setVideoMode 
    
    xor bh, bh                 ;Coloca no BH o valor 0, para a pagina ativa ser a 0  
   
    call writeAskOpen              
               
    inc dh           
    call setCursorPosition     
    xor cx, cx
    
    mov di, offset Filename    
    
    askCharOpen:
    
    call getchar               ;Pede um caracter ao utilizador
    
    cmp al, 0DH                ;Verifica se o caracter e um "enter"
    je fileOpen              
    
    cmp al, 30H                
    jl askCharOpen
    
    cmp al, 3AH                
    jl validCharOpen
    
    cmp al, 41H                
    jl askCharOpen
    
    cmp al, 5BH                
    jl validCharOpen
    
    cmp al, 61H                
    jl askCharOpen
    
    cmp al, 7AH               
    jg askCharOpen 
    
    validCharOpen:
    
    mov [di], al              
    call printChar             
    inc di                     
    inc cl                     ;Incrementa o stringLength
    
    cmp cl, 15
    je fileOpen              
    
    jmp askCharOpen                
    
    fileOpen: 
    
    or cl, cl                  ;Proteccao caso seja posto apenas um enter como nome de ficheiro
    jz askCharOpen 
    
    mov [di], '.'              
    inc di
    mov [di], 't'
    inc di
    mov [di], 'x'
    inc di
    mov [di], 't'
    inc di
    mov [di], 0
    add cl, 4                    
    
    mov dx, offset Filename                         
    mov al, 0                            ;AL com 1 = Modo de abertura de ficheiro : read
    call fopen                           
    mov bx, ax                           ;Coloca na variavel FileHandle, o fileHandle
    
    mov di, offset Filename
    call clearString                     ;Limpa a string para ser usada em saves futuros
                     
    mov dx, offset TemporaryBuffer                 
    mov cx, 110  
    call fread  
    
    mov di, offset TemporaryBuffer
    mov si, offset cells 
    
    xor cx,cx 
    xor bx,bx
    
    selectColumn:
    
    cmp [di], 'A'
    je readA
    
    cmp [di], 'B'
    je readB
    
    cmp [di], 'C'
    je readC 
    
    cmp [di], 'D'
    je readD
    
    cmp [di], 'F'
    je formulaImport
    jmp endImport 
    
    readA:
    
    inc di
    mov bl, [di]
    sub bl, 30h
    mov si, bx
    dec si
    inc di 
    xor cx,cx
    jmp cellSelect
    
    readB:
    
    inc di
    mov bl, [di]
    sub bl, 30h
    mov si, bx
    dec si
    add si, 4
    inc di 
    xor cx,cx
    jmp cellSelect 
    
    readC:
    
    inc di
    mov bl, [di]
    sub bl, 30h
    mov si, bx
    dec si
    add si, 8
    inc di 
    xor cx,cx
    jmp cellSelect
    
    readD:
    
    inc di
    mov bl, [di]
    sub bl, 30h
    mov si, bx
    dec si
    add si, 12
    inc di
    xor cx,cx
    jmp cellSelect
    
    cellSelect: 
    inc di
    ;cmp [di], 02Dh
    ;je putNegative
    inc cx
    cmp [di], ';'
    jne cellSelect 
    
    ;putNegative:
    ;add [si], 128 
    ;mov bh, 1 
    ;inc si
    ;inc cx
    ;inc di  
    ;mov bh, 1
    ;jmp cellSelect
    
    push cx
    dec cx
    dec di
    mov al, [di]
    sub al, 30h
    mov [si], al 
    dec cx
    cmp cx, 0
    je endSoma
    dec di
    mov al, [di]
    sub al, 30h
    mov bl, 10
    mul bl
    add [si], al
    dec cx
    cmp cx, 0
    je endSoma
    dec di
    mov al, [di] 
    sub al, 30h
    mov bl, 100
    mul bl
    add [si], al
    jmp endSoma  
    
    formulaImport:
    push cx
    xor cx, cx
    mov si, offset FormulaReady 
    add di, 4 
    
    repeatFormula:
    mov bl, [di]
    mov [si], bl
    inc di
    inc si
    inc cx 
    cmp cx, 5
    jne repeatFormula
    jmp endImport
   
    endSoma:
    cmp bh, 1
    ;je inverter
    ;keep:
    pop cx
    add di, cx 
    jmp selectColumn 
    
    ;inverter:
    ;neg [si]
    ;jmp keep
    
    endImport:
    
    
    ret   
  importFile endp  
  
;************************************************************
;writeAskOpen
;            
;Description: puts the string into the screen
;Input: 
; ask = string that contains asks the filename
;Output:Nothing
;Destroys: AX, BX, CX, DX, BP
;************************************************************
  
  writeAskOpen proc      ;sujeito a alteracao se apetecer
    
    mov ax, 1301h
    mov bh, 0
    mov bl, 0Fh
    mov cx, 18
    mov dh, 3
    mov dl, 3
    push ds
    pop es
    mov bp, offset AskOpen
    int 10h
    
    ret
  writeAskOpen endp 
  
;************************************************************
;fcreate 
; 
;Description: creats a file in read or write mode
;Input:
; DS:DX - Contains the offset to the file name
;Output:
; CF - 0 if sucessful, 1 if error
; AX - File Handler if sucessful, Error code if error
;Destroys: AX, CX
;************************************************************      
  
  fcreate proc 
               
    mov cx, 0             
    mov ah, 3CH             
    int 21H 
    jc error2
    
    error2:                
    
    ret                             
  fcreate endp 

;************************************************************
;fopen
;
;Description: opens a file in read, write or read/write mode
;Input:
; DS:DX - Contains the offset to the file name
; AL - acess mode
;Output:
; CF - 0 if sucessful, 1 if error
; AX - File Handler if sucessful, Error code if error
;Destroys:AX
;************************************************************ 

  fopen proc 
    
    mov ah, 3DH             
    int 21H
    jc error
    
    error:                     
    
    ret                    
  fopen endp    

;************************************************************
;fclose 
;
;Description: closes a file
;Input:
; BX - File Handle
;Output:
; CF - 0 if sucessful, 1 if error
; AX - Error code if error
;Destroys:AX
;************************************************************ 

  fclose proc
    
    mov ah, 3EH          
    int 21H               
   
    ret                       
  fclose endp

;************************************************************
;fread
;
;Description: reads from file to memory
;Input:
; BX - File Handle
; CX - number of bytes to read
; DS:DX - memory address
;Output:
; CF - 0 if sucessful, 1 if error 
; AX - number of bytes read, error code if error, or 0 if EOF
;Destroys:AX
;************************************************************  

  fread proc
 
    mov ah, 3FH                    
    int 21h                        
    
    ret                               
  fread endp 

;************************************************************
;fwrite
;
;Description - writes in file from memory
;Input:
; BX - File Handle 
; CX - number of bytes to write 
; DS:DX - memory adress
;Output:
; CF - 0 if sucessful, 1 if error 
; AX - number of bytes written or error code if error
;Destroys:AX
;************************************************************
  
  fwrite proc  
    
    mov ah, 40H    
    int 21H        
    
    ret              
  fwrite endp
  
;************************************************************
;getChar
;
;Description: gets a character from keyboard
;Input:Nothing
;Output:
; AL - byte read
;Destroys: AH
;************************************************************

  getChar proc
    
    mov ah, 07H
	  int 21H 
	
	  ret
  getChar endp

;************************************************************
;printChar
;
;Description: writes a character on screen
;Input:Nothing
;Output:Nothing
;Destroys: AH
;************************************************************  
  
  printChar proc
    
    mov ah, 0EH    ;Move para o AH, o valor 0EH
    int 10H        ;Interrupt 21H 
    
    ret              
  printChar endp 
  
;************************************************************
;writeAsk
;        
;Description: writes the string in the screen
;Input: 
; ask = string that contains asks the filename
;Output:Nothing
;Destroys: AX, BX, CX, DX, BP
;************************************************************
  
  writeAsk proc      ;deveria ser otimizada
    
    mov ax, 1301h
    mov bh, 0
    mov bl, 0Fh
    mov cx, 32
    mov dh, 3
    mov dl, 3
    push ds
    pop es
    mov bp, offset Ask
    int 10h
    
    ret
  writeAsk endp 
  
;************************************************************
;clearString
;
;Description: clears the filename string for future saves
;Input:
; DI - offset of filename
;Output:Nothing
;Destroys: AL
;************************************************************
  
  clearString proc
    
    xor al,al
    rep stosb
    
    ret
  clearString endp
  
;**************************************************************************************
;convertNumberASCII
;
;Description: converts a number into ascii form
;Input:
; DI - offset TemporaryBuffer
; SI - offset of the cell
; AL - number in the cell
;Output:Nothing
;Destroys:CX, BX, AX, DI
;***************************************************************************************

  convertNumberASCII proc
    
    mov bl, 10                           ;Move para o BL, o valor 10
    xor cl, cl                           ;Limpa o CX
    
    cmp al, -1                            ;Comparacao para ver se o numero e negativo
    jg ASCIIcycle
    
    neg al                               ;Complemento para dois
    inc bh                               ;BH a 1 = Numero Negativo
    mov [di], '-'                        ;Move para a string, o sinal de negativo
    inc di                               ;Incrementa DI
    
    ASCIIcycle:
   
    cmp ax, 10
    jl resto                             ;Resto da divisao

    div bl                               ;Divisao
    add ah, '0'                          ;Converte para ASCII
    push ax                              ;Guarda um digito
    xor ah, ah                           ;Coloca AH a 0
    inc cl                               ;Insercao de um caracter
    jmp ASCIIcycle 
    
    resto:                               ;Mais um digito
    inc cl                               ;Move para o CH, o CL
    mov ch, cl
    
    add al, '0'                          ;Converte para ASCII
    mov [di], al                         ;Coloca o caracter
    
    writeStringFile: 
    
    cmp ch, 1                            ;Comparacao a ver se todos os digitos ja foram colocados
    je endString 
    
    inc di
    
    pop ax
    mov [di], ah                         ;Move caracter para a string
    
    dec ch                               ;A cada caracter que coloca na string, decrement o ch
    jmp writeStringFile
   
    endString:
                                         
    add cl, bh                           ;Caso o numero seja negativo, mais um caracter
    dec ch                               ;Decremento para o ch ter valor 0
    
    ret                                  ;Return
    
  convertNumberASCII endp
  
;************************************************************
;fWriteCell
;
;Description: writes a cell in the save file
;Input:
; SI - offset cells
;Output:Nothing
;Destroys: Nothing
;************************************************************
  
  fWriteCell proc  
    
    cmp si, 3
    jle A
    
    cmp si, 7
    jle B
    
    cmp si, 11
    jle C
    
    mov [di], 'D'
    jmp cellNumber
    
    A: 
    mov [di], 'A'
    jmp cellNumber
    
    B:
    mov [di], 'B'
    jmp cellNumber
    
    C:
    mov [di], 'C' 
    
    cellNumber:
    
    inc di                               
    
    push ax                              
    mov ax, si                           
    mov bx, 4                            
    div bl                               
    add ah, '1'                          ;Adiciona ao AH, '1', a modo de converter o codigo ASCII. +1
    mov [di], ah                         ;Coloca na string o numero da celula
    pop ax                               ;Recupera o valor de AX
    
    inc di                               
    mov [di], ':'                       
    inc di 
    
    call convertNumberASCII              
    
    add cl, 4                            ;Adiciona ao CL 4, tendo o CL o numero de bytes a escrever
    inc di                               
    mov [di], ';' 
    
    ret                                  
  fWriteCell endp 
  
;************************************************************
;fFormulaCell
;
;Description: loads the formula segment to the export file
;Input: Nothing
;Output:Nothing
;Destroys: DI, CX, BL, SI
;************************************************************
                   
  fFormulaCell proc
    
    push cx
    xor cx, cx
                   
    mov [di], 'F'
    inc di
    mov [di], 'O'
    inc di
    mov [di], 'R'
    inc di
    mov [di], ':'
    inc di             
    mov si, offset formulaReady
    
    formula2:
    mov bl, [si]
    mov [di], bl
    inc di
    inc si
    inc cx
    cmp cx, 5
    jbe formula2
                                  
    retornar:              
    pop cx              
                   
    ret
  fFormulaCell endp
    
    
  
;************************************************************
;exportFile
;
;Description: saves memory in a text file named by the user
;Input:Nothing
;Output:Nothing
;Destroys: CX, DX, BX, AX, SI, DI
;************************************************************
  
   exportFile proc   
     
    startExport: 
   
    mov al, 13H                
    call setVideoMode 
    
    xor bh, bh                 ;Coloca no BH o valor 0, para a pagina ativa ser a 0  
   
    call writeAsk              
               
    inc dh           
    call setCursorPosition     
    xor cx, cx
    
    mov di, offset Filename    
    
    askChar:
    
    call getchar               ;Pede um caracter ao utilizador
    
    cmp al, 0DH                ;Verifica se o caracter e um "enter"
    je fileCreate              ;Cria ficheiro caso seja
    
    cmp al, 30H                
    jl askChar
    
    cmp al, 3AH                
    jl validChar
    
    cmp al, 41H                
    jl askChar
    
    cmp al, 5BH                
    jl validChar
    
    cmp al, 61H                
    jl askChar
    
    cmp al, 7AH               
    jg askChar 
    
    validChar:
    
    mov [di], al              
    call printChar             
    inc di                     
    inc cl                     ;Incrementa o stringLength
    
    cmp cl, 15
    je fileCreate              
    
    jmp askChar                
    
    fileCreate: 
    
    or cl, cl                  ;Proteccao caso seja posto apenas um enter como nome de ficheiro
    jz askChar                 
    
    mov [di], '.'              
    inc di
    mov [di], 't'
    inc di
    mov [di], 'x'
    inc di
    mov [di], 't'
    inc di
    mov [di], 0
    add cl, 4                            ;Incrementa o stringLenght por 4 devido ao .txt
    
    mov dx, offset Filename 
    call fcreate                         
    
    mov al, 1                            ;AL com 1 = Modo de abertura de ficheiro : write
    call fopen                           
    mov bx, ax                           ;Coloca na variavel FileHandle, o fileHandle
    
    mov di, offset Filename
    call clearString                     ;Limpa a string para ser usada em saves futuros
    
    mov si, offset cells                 
    mov dx, offset TemporaryBuffer       
    
    compareValue128: 
    
    xor cx, cx 
    
    mov al, byte ptr [si]                                         
    
    cmp al, 128                          
    je Value128 
    
    WTemporaryBuffer:  
    push bx                              
    mov di, offset TemporaryBuffer       
    call fWriteCell                      
    pop bx                               ;Recupera o valor de BX
    call fwrite                          
    mov di, offset TemporaryBuffer       
    call clearString                     ;Limpa o TemporaryBuffer, para ser usado no futuro
    inc si                               
    
    compare16:
    
    cmp si, 16                           ;Compara para ver se ja atingiu a ultima celula
    jl compareValue128
    
    writeOperando:    
    push bx      
    mov di, offset TemporaryBuffer
    call fFormulaCell
    pop bx  
    mov cx, 9
    call fwrite
    
    call fclose                          
    
    ret                                  
    
    Value128:
                                         
    inc si                                                             
    jmp compare16                     
           
    ret         
  exportFile endp  
  
;************************************************************
;loadContents
;
;Description:at the start of the program checks if there is a contents.bin file and if there is, it is read and written in memory
;Input:Nothing
;Output:Nothing
;Destroys:AX, BX, CX, DX
;************************************************************ 

  loadContents proc
    
    mov dx, offset Contents   ;Passagem para o DX do offset da string que contem "contents.bin"
    xor al, al                ;AL com 0 = Modo de abertura de ficheiro : read
    call fopen                ;Abre o Ficheiro contents.bin caso exista
    jc exitLoadContents       ;Caso o programa nao exista
    mov bx, ax                ;Move para o BX, o file handle
   
    mov cx, 16                ;Numero de bytes a escrever
    mov dx, offset cells      
    call fread
    mov dx, offset formulaReady
    call fread                 
    call fclose               
   
    exitLoadContents:
    
    ret  
   loadContents endp
  
end start ; set entry point and stop the assembler.