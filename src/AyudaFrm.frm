VERSION 5.00
Begin VB.Form AyudaFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   Icon            =   "AyudaFrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   10
      Top             =   6780
      Width           =   7875
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "- Buscar cualquier tarea existente y acceder a su día directamente (se puede Imprimir)."
      Height          =   195
      Index           =   6
      Left            =   480
      TabIndex        =   15
      Top             =   2160
      Width           =   7395
   End
   Begin VB.Label Lbls 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"AyudaFrm.frx":0442
      Height          =   435
      Index           =   12
      Left            =   120
      TabIndex        =   14
      Top             =   5580
      Width           =   7755
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7860
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "- Consultar la agenda de Notas / Teléfonos (se puede Imprimir)."
      Height          =   195
      Index           =   7
      Left            =   480
      TabIndex        =   13
      Top             =   1950
      Width           =   7395
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "- Activar / Desactivar una tarea (doble click en la lista sobre la columna ""Activa"")."
      Height          =   195
      Index           =   14
      Left            =   480
      TabIndex        =   12
      Top             =   1740
      Width           =   7395
   End
   Begin VB.Label Lbls 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"AyudaFrm.frx":0505
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   120
      TabIndex        =   11
      Top             =   6060
      Width           =   7755
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"AyudaFrm.frx":0612
      Height          =   795
      Index           =   11
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   7755
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"AyudaFrm.frx":076B
      Height          =   795
      Index           =   10
      Left            =   120
      TabIndex        =   8
      Top             =   3660
      Width           =   7755
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"AyudaFrm.frx":08EB
      Height          =   615
      Index           =   9
      Left            =   120
      TabIndex        =   7
      Top             =   2940
      Width           =   7755
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"AyudaFrm.frx":0A06
      Height          =   435
      Index           =   8
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   7755
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "- Añadir / Modificar / Borrar nuevo aviso de Tarea (con opción de sonido en formato onda (.WAV))."
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   1548
      Width           =   7395
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "- Salir (y detener)  la Agenda."
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   1344
      Width           =   7395
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "- Activar / Detener la Agenda."
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   1140
      Width           =   7395
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "El funcionamiento es muy sencillo. Las opciones posibles son:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   7755
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Para ello emplea una base de datos creada en Access 97 de un tamaño muy pequeño (80 Kb sin datos)."
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7755
   End
   Begin VB.Label Lbls 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"AyudaFrm.frx":0A9B
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7755
   End
End
Attribute VB_Name = "AyudaFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnSalir_Click()
  Unload Me
End Sub
