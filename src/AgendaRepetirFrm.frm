VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form AgendaRepetirFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Repetir evento"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnAceptar 
      BackColor       =   &H00FFEED9&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton BtnCancelar 
      BackColor       =   &H00FFEED9&
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1635
      Left            =   120
      TabIndex        =   17
      Top             =   540
      Width           =   5655
      Begin VB.TextBox TxtOb 
         BackColor       =   &H00E8FFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "15"
         Top             =   1200
         Width           =   315
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFEED9&
         Caption         =   "El día X de cada mes"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   11
         Top             =   1260
         Width           =   1875
      End
      Begin VB.CheckBox ChkDias 
         BackColor       =   &H00FFEED9&
         Caption         =   "Dom"
         Enabled         =   0   'False
         Height          =   195
         Index           =   6
         Left            =   4800
         TabIndex        =   10
         Top             =   900
         Width           =   675
      End
      Begin VB.CheckBox ChkDias 
         BackColor       =   &H00FFEED9&
         Caption         =   "Sab"
         Enabled         =   0   'False
         Height          =   195
         Index           =   5
         Left            =   4080
         TabIndex        =   9
         Top             =   900
         Width           =   615
      End
      Begin VB.CheckBox ChkDias 
         BackColor       =   &H00FFEED9&
         Caption         =   "Vie"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   3360
         TabIndex        =   8
         Top             =   900
         Width           =   615
      End
      Begin VB.CheckBox ChkDias 
         BackColor       =   &H00FFEED9&
         Caption         =   "Jue"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   7
         Top             =   900
         Width           =   615
      End
      Begin VB.CheckBox ChkDias 
         BackColor       =   &H00FFEED9&
         Caption         =   "Mie"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   6
         Top             =   900
         Width           =   615
      End
      Begin VB.CheckBox ChkDias 
         BackColor       =   &H00FFEED9&
         Caption         =   "Mar"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   900
         Width           =   615
      End
      Begin VB.CheckBox ChkDias 
         BackColor       =   &H00FFEED9&
         Caption         =   "Lun"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   900
         Width           =   615
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFEED9&
         Caption         =   "Los días:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFEED9&
         Caption         =   "Todos los días"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin MSComCtl2.DTPicker DtFecha 
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16384001
      CurrentDate     =   40490
   End
   Begin MSComCtl2.DTPicker DtFecha 
      Height          =   315
      Index           =   1
      Left            =   3540
      TabIndex        =   1
      Top             =   120
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16384001
      CurrentDate     =   40490
   End
   Begin VB.Label LblOb 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFEED9&
      Caption         =   "Hasta"
      Height          =   195
      Index           =   1
      Left            =   3000
      TabIndex        =   16
      Top             =   180
      Width           =   420
   End
   Begin VB.Label LblOb 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFEED9&
      Caption         =   "Desde"
      Height          =   195
      Index           =   0
      Left            =   1020
      TabIndex        =   15
      Top             =   180
      Width           =   465
   End
End
Attribute VB_Name = "AgendaRepetirFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Se le pasan los datos de la tarea
Public mEstrFechaInicial As String
Public mEstrHora As String
Public mEstrTarea As String
Public mEstrSonido As String

'Para saber la opción elegida
Dim intOpt As Integer

'Valores iniciales
Private Sub Form_Load()
  DtFecha(0).Value = CDate(mEstrFechaInicial) + 1
  DtFecha(1).Value = CDate(mEstrFechaInicial) + 30
End Sub

'Preparar todas las cadenas Sql de inserción a ejecutar para la repetición y dejarlas listas en la variable global GarrSql
Private Sub BtnAceptar_Click()
  Dim I As Date, J As Integer
  
  J = 0
  For I = DtFecha(0).Value To DtFecha(1).Value
    Select Case intOpt
      Case 0
        'Todos los días son válidos
      Case 1
        'Solo si conincide con el día de la semana elegido, para ello, si coincide con alguno, seguir, si no coincide con ninguno, saltar
        If Format(I, "w") = 1 And ChkDias(6).Value = vbChecked Then
        ElseIf Format(I, "w") = 2 And ChkDias(0).Value = vbChecked Then
        ElseIf Format(I, "w") = 3 And ChkDias(1).Value = vbChecked Then
        ElseIf Format(I, "w") = 4 And ChkDias(2).Value = vbChecked Then
        ElseIf Format(I, "w") = 5 And ChkDias(3).Value = vbChecked Then
        ElseIf Format(I, "w") = 6 And ChkDias(4).Value = vbChecked Then
        ElseIf Format(I, "w") = 7 And ChkDias(5).Value = vbChecked Then
        Else
          GoTo Saltar
        End If
      Case 2
        'El día debe ser el marcado, sino saltar
        If Format(I, "d") <> CInt(TxtOb(0)) Then GoTo Saltar
    End Select
    
    'Aumentar el array
    ReDim Preserve GarrSql(J)
    GarrSql(J) = "insert into TAREAS (FECHA, HORA, TAREA, SONIDO, ACTIVA, REPETIDA) values ('" & I & "','" & mEstrHora & "','" & mEstrTarea & "','" & mEstrSonido & "','S','S')"
  
    J = J + 1
Saltar:
  Next I
  
  Unload Me
End Sub

'Salir sin repetir
Private Sub BtnCancelar_Click()
  ReDim GarrSql(0)
  GarrSql(0) = ""
  Unload Me
End Sub

'Seleccionar texto
Private Sub TxtOb_GotFocus(Index As Integer)
  With TxtOb(Index)
    .SelStart = 0
    .SelLength = Len(TxtOb(Index))
  End With
  
  If Index = 0 Then
    BtnAceptar.Default = False
  Else
    BtnAceptar.Default = True
  End If
End Sub

'Comprobar teclas pulsadas
Private Sub TxtOb_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case 0
      Select Case KeyAscii
        Case Asc("0") To Asc("9")
          'Se ha pulsado un nº
        Case 8
          'Se ha pulsado BCKSPC
        Case Else
          KeyAscii = 0
          Beep
      End Select
      
      If CInt(TxtOb(Index)) > 31 Then
        MsgBox "El día tiene que estar entre 1 y 31", vbInformation
        TxtOb(Index) = ""
      End If
  
  End Select
End Sub

'Asignar la variable
Private Sub Opt_Click(Index As Integer)
  Dim I As Byte
  
  If Opt(Index).Value = True Then
    intOpt = Index
  End If
  
  Select Case Index
    Case 0
      For I = 0 To 6
        ChkDias(I).Enabled = False
      Next I
      TxtOb(0).Enabled = False
    Case 1
      For I = 0 To 6
        ChkDias(I).Enabled = True
      Next I
      TxtOb(0).Enabled = False
    Case 2
      For I = 0 To 6
        ChkDias(I).Enabled = False
      Next I
      TxtOb(0).Enabled = True
  End Select
End Sub
