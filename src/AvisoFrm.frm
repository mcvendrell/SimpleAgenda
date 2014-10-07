VERSION 5.00
Begin VB.Form AvisoFrm 
   BackColor       =   &H00FFEED9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aviso"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label LblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Prueba de mensaje   Prueba de mensaje   Prueba de mensaje   Prueba de mensaje   Prueba de mensaje   "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6270
   End
End
Attribute VB_Name = "AvisoFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Se le pasa el tiempo que se desea el mensaje en pantalla y el mensaje a mostrar
Public mEintTiempo As Integer
Public mEstrMsg As String

Private Sub Form_Load()
  'Por defecto 5 segundos
  If mEintTiempo = 0 Then mEintTiempo = 5000
  LblMsg.Caption = mEstrMsg
End Sub

Private Sub Form_Activate()
  Dim I As Integer
  
  For I = 1 To 5
    DoEvents
    Sleep mEintTiempo / 5
    
    LblMsg.BackColor = IIf(I Mod 2 = 0, &HFFFFFF, &HC0E0FF)
  Next I
  
  mEintTiempo = 0
  mEstrMsg = 0
  Unload Me
End Sub

