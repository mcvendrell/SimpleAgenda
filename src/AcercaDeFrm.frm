VERSION 5.00
Begin VB.Form AcercaDeFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de..."
   ClientHeight    =   3690
   ClientLeft      =   30
   ClientTop       =   285
   ClientWidth     =   4995
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnAceptar 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   348
      Left            =   60
      TabIndex        =   0
      Top             =   3300
      Width           =   4905
   End
   Begin VB.PictureBox PicIcono 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "AcercaDeFrm.frx":0000
      ScaleHeight     =   323.838
      ScaleMode       =   0  'User
      ScaleWidth      =   323.838
      TabIndex        =   1
      Top             =   168
      Width           =   480
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFEED9&
      Caption         =   "Web"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2580
      TabIndex        =   13
      Top             =   2700
      Width           =   405
   End
   Begin VB.Label LblWeb 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "http://manuel.conde.name"
      Height          =   195
      Left            =   2850
      TabIndex        =   12
      Top             =   2955
      Width           =   1905
   End
   Begin VB.Label LblEmail 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "manu_conde@hotmail.com"
      Height          =   195
      Left            =   540
      TabIndex        =   11
      Top             =   2955
      Width           =   1935
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFEED9&
      Caption         =   "e-Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   270
      TabIndex        =   10
      Top             =   2700
      Width           =   525
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFEED9&
      Caption         =   "Desarrollada por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   870
      TabIndex        =   5
      Top             =   2100
      Width           =   1410
   End
   Begin VB.Label LblDescripcion 
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1140
      TabIndex        =   9
      Top             =   1635
      Width           =   3750
   End
   Begin VB.Label LblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1140
      TabIndex        =   8
      Top             =   1035
      Width           =   45
   End
   Begin VB.Label LblAutor 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manuel Conde Vendrell"
      Height          =   195
      Left            =   1140
      TabIndex        =   7
      Top             =   2385
      Width           =   1650
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1140
      TabIndex        =   6
      Top             =   420
      Width           =   45
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFEED9&
      Caption         =   "Versión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   870
      TabIndex        =   4
      Top             =   768
      Width           =   645
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFEED9&
      Caption         =   "Título de la aplicación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   870
      TabIndex        =   3
      Top             =   165
      Width           =   1935
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFEED9&
      Caption         =   "Descripción de la aplicación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   870
      TabIndex        =   2
      Top             =   1371
      Width           =   2430
   End
End
Attribute VB_Name = "AcercaDeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Se carga el form
Private Sub Form_Load()
  Me.Caption = "Acerca de " & App.Title
  LblTitulo.Caption = App.Title
  LblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
  LblDescripcion.Caption = App.FileDescription
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Al salir de la label, restablecer su color
  If LblEmail.FontUnderline Then
    LblEmail.FontUnderline = False
    LblEmail.ForeColor = &H0
  End If

  If LblWeb.FontUnderline Then
    LblWeb.FontUnderline = False
    LblWeb.ForeColor = &H0
  End If
End Sub

'Se pulsó aceptar
Private Sub BtnAceptar_Click()
  Unload Me
End Sub

'Lanzar el mail
Private Sub LblEmail_Click()
  ShellExecute hwnd, "open", "mailto:manuel_conde@hotmail.com", vbNullString, vbNullString, SW_SHOW
End Sub

'Destacar el enlace
Private Sub LblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not LblEmail.FontUnderline Then
    LblEmail.FontUnderline = True
    LblEmail.ForeColor = &HD96D44
  End If
End Sub

'Lanzar la web
Private Sub LblWeb_Click()
  ShellExecute hwnd, "open", "http://manuel.conde.name", vbNullString, vbNullString, SW_SHOW
End Sub

'Destacar el enlace
Private Sub LblWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not LblWeb.FontUnderline Then
    LblWeb.FontUnderline = True
    LblWeb.ForeColor = &HD96D44
  End If
End Sub

