VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin Proyecto1.Clista Clista1 
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2175
      _extentx        =   3836
      _extenty        =   2566
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer

Private Sub Command1_Click()
    n = n + 1
    Clista1.Agregar n & ": fila" & vbTab & "un lugar" & vbTab & "de la mancha", vbRed
    
End Sub

Private Sub Command2_Click()
    Clista1.AgregarCol "COLUM" & Clista1.nCol + 1
End Sub

Private Sub Form_Load()
    Clista1.AgregarCol "COLUM1"
    Clista1.AgregarCol "COLUM2"
    Clista1.Agregar "En" & vbTab & "un lugar" & vbTab & "de la mancha"
    
End Sub

Private Sub Form_Resize()
    Clista1.Left = 0
    Clista1.Width = Me.Width * 0.95
    Clista1.Height = Me.Height * 0.7
    Clista1.Refrescar
End Sub

'Private Sub List1_Click()
'    Debug.Print List1.listIndex
'End Sub

Private Sub Clista1_Click()
    Debug.Print Clista1.listIndex
End Sub


