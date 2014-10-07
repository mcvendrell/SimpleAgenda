VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form TlfnsFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda Telefónica"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   Icon            =   "TlfnsFrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Introduzca los valores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   60
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00FFEED9&
         Caption         =   "Cancelar"
         Height          =   360
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   260
         Width           =   2175
      End
      Begin VB.CommandButton BtnAceptar 
         BackColor       =   &H00FFEED9&
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   360
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   260
         Width           =   2175
      End
      Begin VB.TextBox TxtOb 
         BackColor       =   &H00E8FFFF&
         Height          =   300
         Index           =   1
         Left            =   840
         MaxLength       =   100
         TabIndex        =   6
         Text            =   "Escriba aquí lo que quiera (100 caracteres máximo)"
         Top             =   660
         Width           =   8055
      End
      Begin VB.TextBox TxtOb 
         BackColor       =   &H00E8FFFF&
         Height          =   300
         Index           =   0
         Left            =   840
         MaxLength       =   25
         TabIndex        =   5
         Text            =   "Nombre (25 caracteres)"
         Top             =   300
         Width           =   2535
      End
      Begin VB.Label LblOb 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFEED9&
         Caption         =   "Tlfn"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   705
         Width           =   270
      End
      Begin VB.Label LblOb 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFEED9&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   345
         Width           =   555
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlxDatos 
      Height          =   3195
      Left            =   60
      TabIndex        =   9
      Top             =   2160
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   5636
      _Version        =   393216
      BackColor       =   16777215
      BackColorSel    =   14221054
      ForeColorSel    =   0
      BackColorBkg    =   16772825
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   $"TlfnsFrm.frx":0442
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox PicBotones 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   9000
      TabIndex        =   10
      Top             =   120
      Width           =   9000
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00FFEED9&
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1755
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00FFEED9&
         Caption         =   "&Modificar Registro"
         Height          =   375
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1755
      End
      Begin VB.CommandButton BtnBorrar 
         BackColor       =   &H00FFEED9&
         Caption         =   "&Eliminar Registro"
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1755
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00FFEED9&
         Caption         =   "&Añadir Registro"
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   0
         Width           =   1755
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00FFEED9&
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1755
      End
   End
   Begin VB.Label Lbls 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datos existentes"
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
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   1860
      Width           =   9075
   End
End
Attribute VB_Name = "TlfnsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Aqui tenemos una conexión independiente de la del otro formulario, por lo que no
'hay problemas al salir se cierra esta conexión

Dim Sql As String
Dim BD As Database

Dim mblnModificando As Boolean

'Conecta con la BD
Private Function ConectaBD() As Boolean
  On Local Error Resume Next
  
  Set BD = OpenDatabase(App.Path & "\Agenda.mdb")
  
  If Err.Number <> 0 Then
    MsgBox "Se produjo un error al intentar abrir la Base de Datos: " & Err.Description
    MsgBox "Asegúrese de que el archivo 'Agenda.mdb' se encuentra en el directorio " & vbCrLf & vbCrLf & App.Path
    ConectaBD = False
  Else
    ConectaBD = True
  End If
End Function

'Conecta con la BD
Private Sub DesconectaBD()
  'Liquidar la conexión para que no se empleen recursos del sistema
  On Local Error Resume Next
  BD.Close
  Set BD = Nothing
End Sub

'Resetea el flex de datos
Private Sub InicializaFlex()
  FlxDatos.Clear
  FlxDatos.FormatString = "<Nombre                                       |<Tlfn - Descripción                                                                                                                "
  FlxDatos.Rows = 2
End Sub

'Rellena el flex de datos
Private Sub RefrescaFlex()
  Dim Rec As Recordset
  Dim I As Integer
  
  Screen.MousePointer = vbHourglass
  
  Sql = "select NOMBRE, DATOS" _
      & "  from TLFNS" _
      & " order by NOMBRE"
  Set Rec = BD.OpenRecordset(Sql)
  
  InicializaFlex
  I = 0
  While Not Rec.EOF
    I = I + 1
    
    FlxDatos.TextMatrix(I, 0) = Rec!NOMBRE
    FlxDatos.TextMatrix(I, 1) = Rec!DATOS
    
    Rec.MoveNext
    
    'Si va a haber mas filas, añadir una row más
    If Not Rec.EOF Then FlxDatos.Rows = FlxDatos.Rows + 1
  Wend
  Rec.Close

  If FlxDatos.Visible And FlxDatos.Enabled Then FlxDatos.SetFocus
  FlxDatos.Row = 1
  
  Set Rec = Nothing
  Screen.MousePointer = vbDefault
End Sub

'Cambiar muestra
Private Sub FlxDatos_DblClick()
  On Error GoTo Errores
  
  If FlxDatos.TextMatrix(1, 0) <> "" Then
    Call BtnModificar_Click
  Else
    MsgBox "No hay datos.", vbInformation
  End If
  Exit Sub

Errores:
  MsgBox "Error: " & Err.Description
End Sub

'Conectarse a la BD y leer datos
Private Sub Form_Load()
  If ConectaBD Then RefrescaFlex
End Sub

'Foco en los botones
Private Sub Form_Activate()
  If PicBotones.Enabled And PicBotones.Visible Then PicBotones.SetFocus
End Sub

'Se descarga el form, desconectar
Private Sub Form_Unload(Cancel As Integer)
  DesconectaBD
End Sub

'Imprimir la BD
Private Sub BtnImprimir_Click()
  Dim Sql As String, Rec As Recordset
  
  Screen.MousePointer = vbHourglass
  
  Sql = "select NOMBRE, DATOS" _
      & "  from TLFNS" _
      & " order by NOMBRE"
  Set Rec = BD.OpenRecordset(Sql)
  
  If Not Rec.EOF Then ImpresionTlfns Rec, False
  
  Rec.Close
  Set Rec = Nothing
  Screen.MousePointer = vbDefault
End Sub

'Terminar la aplicación
Private Sub BtnSalir_Click()
  Unload Me
End Sub

'Añadir una tarea
Private Sub BtnAñadir_Click()
  FraDatos.Visible = True
  mblnModificando = False
  PicBotones.Enabled = False
  FlxDatos.Enabled = False
End Sub

'Borrar tarea seleccionada
Private Sub BtnBorrar_Click()
  On Error GoTo Errores
  
  If FlxDatos.TextMatrix(1, 0) <> "" Then
    If MsgBox("¿Seguro que desea eliminar este registro?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then Exit Sub
    
    Sql = "delete from TLFNS" _
        & " where NOMBRE = '" & FlxDatos.TextMatrix(FlxDatos.Row, 0) & "'"
    BD.Execute Sql
    
    RefrescaFlex
  Else
    MsgBox "No hay datos.", vbInformation
  End If
  Exit Sub

Errores:
  MsgBox "Error: " & Err.Description
End Sub

'Modificar tarea seleccionada
Private Sub BtnModificar_Click()
  If FlxDatos.TextMatrix(1, 0) <> "" Then
    TxtOb(0) = FlxDatos.TextMatrix(FlxDatos.Row, 0)
    TxtOb(1) = FlxDatos.TextMatrix(FlxDatos.Row, 1)
    FraDatos.Visible = True
    mblnModificando = True
    PicBotones.Enabled = False
    FlxDatos.Enabled = False
  Else
    MsgBox "No hay datos.", vbInformation
  End If
End Sub

'Añadir la tarea
Private Sub BtnAceptar_Click()
  On Error GoTo Errores
  
  If TxtOb(0) <> "" And TxtOb(1) <> "" Then
    If mblnModificando Then
      Sql = "update TLFNS" _
          & "   set NOMBRE = '" & TxtOb(0) & "'" _
          & "     , DATOS = '" & TxtOb(1) & "'" _
          & " where NOMBRE = '" & FlxDatos.TextMatrix(FlxDatos.Row, 0) & "'"
    Else
      Sql = "insert into TLFNS" _
          & "       (NOMBRE, DATOS) " _
          & "values ('" & TxtOb(0) & "','" & TxtOb(1) & "')"
    End If
    BD.Execute Sql
    
    RefrescaFlex
    
    PicBotones.Enabled = True
    FlxDatos.Enabled = True
    FraDatos.Visible = False
  Else
    MsgBox "Rellene todos los campos.", vbInformation
  End If
  Exit Sub

Errores:
  MsgBox "Error: " & Err.Description
End Sub

'Cancelar añadir/modificar
Private Sub BtnCancelar_Click()
  FraDatos.Visible = False
  FlxDatos.Enabled = True
  PicBotones.Enabled = True
  PicBotones.SetFocus
End Sub

'Seleccionar texto
Private Sub TxtOb_GotFocus(Index As Integer)
  With TxtOb(Index)
    .SelStart = 0
    .SelLength = Len(TxtOb(Index))
  End With
End Sub

'Comprobar teclas pulsadas
Private Sub TxtOb_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case 0, 1
      'Valen todas menos el apóstrofe
      Select Case KeyAscii
        Case Asc("'")
          'Se ha pulsado un separador
          KeyAscii = 0
          MsgBox "El apóstrofe (') no es un caracter válido. En su lugar puede emplear un acento si lo desea.", vbInformation
      End Select
    
  End Select
End Sub
