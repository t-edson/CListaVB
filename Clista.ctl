VERSION 5.00
Begin VB.UserControl Clista 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picExcel 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3240
      Picture         =   "Clista.ctx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      ToolTipText     =   "Exportar a Excel"
      Top             =   2760
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   840
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   2760
      Width           =   2415
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Left            =   3240
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "Clista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'                          Control CLista
'Implementa un control de lista con varias columnas y opciones de
'color por fila.
'
'                                  por Tito Hinostroza 22/08/2008

Option Explicit

Const ALT_ENCAB_DEF = 250   'Alto del encabezado por defecto
Const ALT_FILA_DEF = 250    'Alto de las filas por defecto
Const ANC_COL_MIN = 50     'Ancho mínimo de columna
Private Mlinvis As Integer  'Máximo número de filas visibles
Private MlinvisC As Integer 'Máximo número de filas visibles completas
Private nlinvis As Integer  'número de filas visibles
Private nlinvisC As Integer 'número de filas visibles completas
Private nFilIni As Integer  'número de fila que empieza a visualizarse

Private fh As Single       'factor de ampliación horizontal
Private fv As Single       'factor de ampliación vertical
Private altfil As Single   'alto de fila
Private dx0 As Single      'desplazamiento anterior de x0

Private col() As Tcolum    'Propiedades de las columnas (ancho, alineamiento, etc)
Public nCol As String      'número de columnas
Private encab As Tfila     'encabezado
Private fil() As Tfila     'objeto filas (cadena, color, etc)
Public nFil As Long        'número de filas
Private despX As Single    'desplazamiento horizontal (<=0)

Private ancTotCol As Single    'ancho total de las columnas
Private altTotFil As Single    'alto total de las filas
Private t As Integer
Private colAdim As Integer  'columna a dimensionar
Private xinAdim As Single   'x inicial para dimensionar columna
Private fil_selec As Long   'fila seleccionada

Public Event Click()        'Para que se comporte como un "ListBox"

Public Function listIndex() As Long
'Propiedad compatible con la de un "ListBox"
    listIndex = fil_selec - 1
End Function

Public Function AddItem(item As String)
'Propiedad compatible con la de un "ListBox"
    Call Agregar(item)
End Function

Public Sub Clear()
'Propiedad compatible con la de un "ListBox"
    UserControl_Initialize
    Call Refrescar
End Sub

Private Sub HScroll1_Change()
Dim i As Integer
Dim f As Single
Dim maxDespX As Single
    maxDespX = -(ancTotCol - (ScaleWidth - 4 * t))
    f = HScroll1.Value / HScroll1.Max
    despX = f * maxDespX
    Call Refrescar
End Sub

Private Sub UserControl_Initialize()
    ReDim col(0)
    nCol = 0
    ancTotCol = 0   'el acho inicial es 0
    ReDim fil(0)
    nFil = 0
    'tamaño del pixel
    t = Screen.TwipsPerPixelX
    
    nFilIni = 1     'empieza a visualizar en la fila 1
    despX = 0
End Sub

Public Sub Refrescar()
'Redibuja en pantalla
Dim f As Long, c As Integer
Dim posY As Single, posX As Single
Dim a() As String
Dim alt_vis As Single   'alto de área visible
Dim linAvis As Integer
Dim fila As Tfila
    If nCol = 0 Then    'protección
        AgregarCol "COLUMNA 1", 1200
    End If
    'líneas totales a visualizar si el control fuera infinitamente largo
    linAvis = nFil
    linAvis = linAvis - nFilIni + 1 'corrección por desplazamiento
    'Calcula líneas visibles en el control
    'quitando el área de la barra horiz. inferior
    alt_vis = ScaleHeight - HScroll1.Height - 2 * t
    'calcula el máximo número de líneas visibles en el control incluyendo
    'el encabezado. Sirve para determinar cuantas líneas dibujar
    Mlinvis = Int(alt_vis / ALT_FILA_DEF) + 1
    'cantidad real de líneas visibles en el control
    nlinvis = Mlinvis
    If nlinvis > linAvis Then nlinvis = linAvis
    'calcula la cantidad de líneas completas que se pueden mostrar.
    'Indica la última línea que puede tener el enfoque
    'antes de iniciar el desplazamiento.
    MlinvisC = Int(alt_vis / ALT_FILA_DEF)
    nlinvisC = MlinvisC
    If nlinvisC > linAvis Then nlinvisC = linAvis
    'borra espacio de dibujo
    Line (0, 0)-(ScaleWidth, ScaleHeight), RGB(204, 204, 204), BF
    'Dibuja datos de encabezado
    posX = despX + t    'posición inicial X (no toca el borde)
    ForeColor = vbBlack 'color de texto de encabezado
    For c = 1 To nCol
        posY = t
'        SetTextAlign hdc, TA_CENTER
'        Line (posX, posY)-(posX + col(c).ancho, posY + ALT_ENCAB_DEF), , B
        BordeSaliente posX + t, posY + t, posX + col(c).ancho, posY + ALT_ENCAB_DEF + t  'Dibuja borde
        MultiTexto hdc, col(c).encab, posX + 2 * t, posY + 2 * t, posX + col(c).ancho, posY + ALT_ENCAB_DEF
        posX = posX + col(c).ancho
    Next
    'Dibuja datos de filas
'    DrawWidth = 1
    posY = ALT_ENCAB_DEF + t  'deja espacio para encabezado
    For f = 1 To nlinvis   'nFil
        fila = fil(f + nFilIni - 1)
        a = Split(fila.txt, vbTab)
        posX = despX + t    'posición inicial X (no toca el borde)
        For c = 1 To nCol
            SetTextAlign hdc, TA_LEFT
            'dibuja borde
            If fila.selecc Then
                Line (posX, posY)-(posX + col(c).ancho, posY + ALT_FILA_DEF), RGB(0, 0, 128), BF
            Else
                Line (posX, posY)-(posX + col(c).ancho, posY + ALT_FILA_DEF), vbWhite, BF
            End If
            Line (posX, posY)-(posX + col(c).ancho, posY + ALT_FILA_DEF), vbBlack, B
            'selecciona color de texto acuerdo a selección
            If fila.selecc Then ForeColor = vbWhite Else ForeColor = fila.txtcol
            'dibuja texto
            MultiTexto hdc, a(c - 1), posX, posY, posX + col(c).ancho, posY + ALT_FILA_DEF
            'Refresh
            'actualiza nueva posición vertical
            posX = posX + col(c).ancho
        Next
        posY = posY + ALT_FILA_DEF
    Next
    'Dibuja Borde
'    ancTotCol = posY    'actualiza ancho total
    altTotFil = posY
    
    BordeHundido 0, 0, Width, Height 'Dibuja borde
        
    Call Resize
End Sub

Public Sub AgregarCol(encab As String, Optional ancho As Single = 800, _
                      Optional alineam As Long = AL_IZQ)
'Agrega una columna al control. Se debe hacer antes de
'agregar las filas
Dim n As Long
    'Refresca y Actualiza Barras de Desplazamiento
    nCol = nCol + 1
    ReDim Preserve col(nCol)
    col(nCol).encab = encab
    col(nCol).ancho = ancho
    col(nCol).alineam = alineam
    ancTotCol = ancTotCol + ancho   'actualiza el ancho total
    'actualiza los datos que pudiera contener
    For n = 1 To nFil
        'completa las tabulaciones adicionales
        CompletaTabs fil(n).txt
    Next
    Call Refrescar
End Sub


Public Function Agregar(txt As String, Optional coltext As Long = vbBlack, _
                                   Optional colfond As Long = vbWhite) As Integer
'Agrega un elemento a la lista. Para ver el control actualizado debe llamarse a
    ' refrescar()
    CompletaTabs txt    'completa tabulaciones
    nFil = nFil + 1
    ReDim Preserve fil(nFil)
    fil(nFil).txt = txt
    fil(nFil).txtcol = coltext
    fil(nFil).foncol = colfond
    Agregar = nFil
    Call Refrescar
End Function

Public Sub Resize()
'Actualiza parámetros de las barras de desplazamiento
'    Call Refrescar
    'posiciona botón de Excel
    picExcel.Left = ScaleWidth - picExcel.Width - t
    picExcel.Top = ScaleHeight - picExcel.Height - t
    'posiciona barra horizontal
    HScroll1.Top = ScaleHeight - HScroll1.Height - t
    HScroll1.Left = 2 * t
    HScroll1.Width = ScaleWidth - 2 * t - picExcel.Width
    HScroll1.Visible = False
    HScroll1.Visible = True 'para forzar a refrescar, sino se dibuja mal
    '-----Inicia barra de desplazamiento horizontal-------
    'Calcula factor de reducción para desplazamiento horizontal
    fh = ScaleWidth / ancTotCol     'factor de reducción
    If fh >= 1 Then  'no es necesario
        HScroll1.Enabled = False
    Else
        HScroll1.Enabled = True
        HScroll1.Min = 0
        HScroll1.Max = 100 'aumenta por factor de 100
'        HScroll1.LargeChange = 100 * fh   'factor que debería ser
        HScroll1.LargeChange = Exp(9 * fh) 'función más cercana al efecto deseado
        HScroll1.SmallChange = 10       'desplaza una parte
    End If
    '-------Inicia barra de desplazamiento vertical--------
    fv = MlinvisC / (UBound(fil) + 1)  'factor de reducción
    'verifica necesidad de existencia de barra vertical
    If fv < 1 Then 'hay más líneas de las que se puede mostrar
        'posiciona barra vertical
        VScroll1.Top = t
        VScroll1.Left = ScaleWidth - VScroll1.Width - t
        VScroll1.Height = Abs(ScaleHeight - picExcel.Height - t)
        'inicia parámetros de desplazamiento
        VScroll1.Min = 0
        VScroll1.Max = UBound(fil) - MlinvisC + 1
        If nlinvisC < 1 Then
            VScroll1.LargeChange = 1
        Else
            VScroll1.LargeChange = nlinvisC
        End If
        VScroll1.SmallChange = 1
        VScroll1.Visible = False    'para forzar a refrescar, sino se dibuja mal
        VScroll1.Visible = True
    Else
        VScroll1.Visible = False
        nFilIni = 1  'para que aprezcan las líneas a partir de la línea 1
        VScroll1.Value = VScroll1.Min
    End If
End Sub

Private Sub MultiTexto(hdc As Long, cad As String, _
        x1 As Single, y1 As Single, x2 As Single, y2 As Single, _
        Optional color As Long = vbBlack)
'Escribe texto en varias líneas, en el rectángulo indicado. Realiza saltos entre palabras
'o cuando una palabra excede el ancho del cuadro. Si una línea adicional no entra completa
'no se visualiza. Si la cadena completa no entra en el cuadro se recorta y se le agrega "..."
Dim x1c As Single, y1c As Single   'coordenadas corregidas
Dim x2c As Single, y2c As Single
Dim rc As RECT
    rc.Left = x1 / Screen.TwipsPerPixelX: rc.Top = y1 / Screen.TwipsPerPixelY
    rc.Right = x2 / Screen.TwipsPerPixelX: rc.Bottom = y2 / Screen.TwipsPerPixelY
'    SetTextColor hDC, color
    DrawText hdc, cad, Len(cad), rc, DT_WORDBREAK + DT_EDITCONTROL + DT_END_ELLIPSIS
End Sub

Private Sub BordeHundido(x1 As Single, y1 As Single, x2 As Single, y2 As Single)
'Dibuja un contorno Hundido, al estilo Windows
    Line (x1, y1)-(x1, y2), RGB(128, 128, 128)  'barra izquierda
    Line (x1, y1)-(x2, y1), RGB(128, 128, 128)  'barra arriba
    
    Line (x2 - t, y1)-(x2 - t, y2), vbWhite     'barra derecha
    Line (x1, y2 - t)-(x2, y2 - t), vbWhite     'barra abajo
    
    Line (x1 + t, y1 + t)-(x1 + t, y2 - 2 * t), RGB(64, 64, 64) 'barra izquierda2
    Line (x1 + t, y1 + t)-(x2 - 2 * t, y1 + t), RGB(64, 64, 64) 'barra arriba2

    Line (x2 - 2 * t, y1)-(x2 - 2 * t, y2), RGB(192, 192, 192) 'barra derecha2
    Line (x1 + t, y2 - 2 * t)-(x2 - 2 * t, y2 - 2 * t), RGB(192, 192, 192) 'barra abajo2
End Sub

Private Sub BordeSaliente(x1 As Single, y1 As Single, x2 As Single, y2 As Single)
'Dibuja un contorno Hundido, al estilo Windows
    Line (x1, y1)-(x1, y2), vbWhite     'barra izquierda
    Line (x1, y1)-(x2, y1), vbWhite     'barra arriba
    
    Line (x2 - t, y1)-(x2 - t, y2), RGB(64, 64, 64) 'barra derecha
    Line (x1, y2 - t)-(x2, y2 - t), RGB(64, 64, 64) 'barra abajo
    
    Line (x1 + t, y1 + t)-(x1 + t, y2 - 2 * t), RGB(192, 192, 192)  'barra izquierda2
    Line (x1 + t, y1 + t)-(x2 - 2 * t, y1 + t), RGB(192, 192, 192)  'barra arriba2

    Line (x2 - 2 * t, y1)-(x2 - 2 * t, y2), RGB(128, 128, 128) 'barra derecha2
    Line (x1 + t, y2 - 2 * t)-(x2 - 2 * t, y2 - 2 * t), RGB(128, 128, 128) 'barra abajo2
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim c As Integer, f As Integer
Dim posY As Single, posX As Single
    If Button = 2 Then Exit Sub
    'botón izquierdo
    If nCol < 1 Then Exit Sub
    posX = despX + t   'posición inicial X (no toca el borde)
    posY = 0
    For c = 1 To nCol
        If X > posX And X < posX + col(c).ancho And Y < ALT_ENCAB_DEF And _
           UserControl.MousePointer <> vbSizeWE Then
            'pulsa el botón de encabezado
            BordeHundido posX + t, posY + t, posX + col(c).ancho, posY + ALT_ENCAB_DEF + t  'Dibuja borde
            Exit Sub
        End If
        posX = posX + col(c).ancho
    Next
    'Si llega aquí es porque no se seleccionó el enzabezado
    
    'quita selección de filas
    If Shift = 0 Then   'sólo si no hay Ctrl o Shift pulsado
        'este es el único lazo que explora toda la lista
        For f = 1 To nFil
            fil(f).selecc = False
        Next
    End If
    'verificar selección de filas
    fil_selec = 0  'sin seleción
    posY = ALT_ENCAB_DEF + t  'deja espacio para encabezado
    For f = 1 To nlinvis   'nFil
        If X < ancTotCol And Y > posY And Y < posY + ALT_FILA_DEF Then
            'lo selecciona
            fil(f + nFilIni - 1).selecc = True
            fil_selec = f + nFilIni - 1      'corrige para que empiece en 0
        End If
        posY = posY + ALT_FILA_DEF
    Next
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim c As Integer
Dim posY As Single, posX As Single
    If nCol < 1 Then Exit Sub
    posX = despX + t   'posición inicial X (no toca el borde)
    If Button = 1 Then
        'se está re-dimensionando el ancho de columna
'        Line (X, 0)-(X, ScaleHeight)
    Else
        'sólo se cambia el puntero si no se ha pulsado el botón izquierdo,
        'esto es para implementar el dimensionamiento de ancho de columna
        UserControl.MousePointer = vbDefault
    End If
    For c = 1 To nCol
        posX = posX + col(c).ancho
        If Abs(posX - X) < 4 * t And Y < altTotFil Then
            'cambia forma de puntero
            UserControl.MousePointer = vbSizeWE
            colAdim = c     'guarda columna a dimensionar
            xinAdim = X     'guarda X inicial
            Exit For
        End If
    Next
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dx As Single
Dim anc_final As Single
Dim c As Integer
    If UserControl.MousePointer = vbSizeWE Then
        'hay que re-dimensionar ancho
        dx = (X - xinAdim)  'desplazamiento
        anc_final = col(colAdim).ancho + dx 'ancho pretendido
        If anc_final > 0 Then
            col(colAdim).ancho = anc_final
        End If
        'actualiza "ancTotCol"
        ancTotCol = 0
        For c = 1 To nCol
            ancTotCol = ancTotCol + col(c).ancho
        Next
    End If
    Call Refrescar  'redibuja para actualizar cambios
    RaiseEvent Click    'dispara evento
End Sub

Private Sub VScroll1_Change()
    nFilIni = VScroll1.Value + 1
    Call Refrescar
End Sub

Private Sub CompletaTabs(txt As String)
'Completa las tabulaciones requeridas en una fila de datos para
'poder mostrarse en el control de lista
Dim n As Integer
Dim a() As String
    a = Split(txt, vbTab)   'separa campos
    For n = UBound(a) + 1 To nCol - 1
        txt = txt & vbTab
    Next
End Sub

Private Sub cmdExcel_Click()
'Exporta los datso de la lista a un formato de Excel e intenta abrir el Excel
Dim nar As Integer
Dim nomarc As String    'nombre de archivo
Dim txt As String
Dim i As Long
Dim c As Integer
    'agrupa datos en variable "txt"
    For c = 1 To nCol
        txt = txt & col(c).encab & vbTab
    Next
    'quita "tab" final
    If Right(txt, 1) = vbTab Then txt = Mid$(txt, 1, Len(txt) - 1)
    txt = txt & vbCrLf
    For i = 1 To nFil
        txt = txt & fil(i).txt & vbCrLf
    Next
    'Crea archivo temporal para exportar a Excel
    nar = FreeFile
    nomarc = "tmp" & Format(Date + Time, "yymmddhhmmss") & ".xls"
    On Error GoTo error_arc
    Open App.Path & "\" & nomarc For Output As #nar
    Print #nar, txt
    Close #nar
    On Error GoTo 0
    On Error GoTo error_xls
    'Nos movemos primero a la ruta porque "start" no trabaja bien con todos los caminos
    If App.Path Like "?:*" Then ChDrive Left(App.Path, 1)
    ChDir App.Path
    Shell "cmd /c start " & nomarc
   
    On Error GoTo 0
    Exit Sub
error_arc:
    MsgBox "Error al exportar. No se puede crear archivo"
    On Error GoTo 0
    Exit Sub
error_xls:
    MsgBox "Error al abrir Excel"
    On Error GoTo 0
End Sub

Private Sub picExcel_Click()
    Call cmdExcel_Click
End Sub
