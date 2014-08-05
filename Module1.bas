Attribute VB_Name = "Module1"
Option Explicit
'-----------Definiciones para la lista------------------
'Propiedades de una fila
Public Type Tfila
    txt As String
    txtcol As Long      'color de texto
    foncol As Long      'color de fondo
    parpad As Boolean   'bandera de parpadeo
    encend As Boolean   'bandera para parpadeo
    selecc As Boolean   'bandera que indica fila seleccionada
End Type

'Propiedades de una columna
Public Type Tcolum
    encab As String     'encabezado de columna
    ancho As Single     'ancho de columna
    alineam As Long     'alineamiento
End Type

'Constantes de alineamiento
Public Const AL_IZQ = 0
Public Const AL_CEN = 1
Public Const AL_DER = 2

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type SIZE
    cx As Long
    cy As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type

Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
  
Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2
Public Const TA_CENTER = 6
Public Const TA_TOP = 0
Public Const TA_BOTTOM = 8
Public Const TA_BASELINE = 24

Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Const DT_BOTTOM = &H8
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_CALCRECT = &H400
Public Const DT_WORDBREAK = &H10
Public Const DT_VCENTER = &H4
Public Const DT_TOP = &H0
Public Const DT_TABSTOP = &H80
Public Const DT_SINGLELINE = &H20
Public Const DT_RIGHT = &H2
Public Const DT_NOCLIP = &H100
Public Const DT_INTERNAL = &H1000
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_EXPANDTABS = &H40
Public Const DT_CHARSTREAM = 4 ' Character-stream, PLP
Public Const DT_NOPREFIX = &H800
Public Const DT_EDITCONTROL = &H2000&
Public Const DT_PATH_ELLIPSIS = &H4000&
Public Const DT_END_ELLIPSIS = &H8000&
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_RTLREADING = &H20000
Public Const DT_WORD_ELLIPSIS = &H40000

Public Const DT_DISPFILE = 6 ' Display-file
Public Const DT_METAFILE = 5 ' Metafile, VDM
Public Const DT_PLOTTER = 0 ' Vector plotter
Public Const DT_RASCAMERA = 3 ' Raster camera
Public Const DT_RASDISPLAY = 1 ' Raster display
Public Const DT_RASPRINTER = 2 ' Raster printer


