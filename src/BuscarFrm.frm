VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form BuscarFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   13485
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnImprimir 
      BackColor       =   &H00FFEED9&
      Caption         =   "&Imprimir"
      Height          =   420
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   675
      Width           =   2775
   End
   Begin VB.CommandButton BtnSalir 
      BackColor       =   &H00FFEED9&
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   420
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1140
      Width           =   2775
   End
   Begin VB.CommandButton BtnAceptar 
      BackColor       =   &H00FFEED9&
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   420
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   225
      Width           =   2775
   End
   Begin VB.Frame FraDatos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Introduzca los valores para la búsqueda. Nada = mostrar todo."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6195
      Begin VB.OptionButton OptBuscar 
         BackColor       =   &H00FFFFFF&
         Caption         =   ">"
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   15
         Top             =   340
         Value           =   -1  'True
         Width           =   435
      End
      Begin VB.OptionButton OptBuscar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "="
         Height          =   195
         Index           =   1
         Left            =   2580
         TabIndex        =   14
         Top             =   340
         Width           =   435
      End
      Begin VB.OptionButton OptBuscar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "<"
         Height          =   195
         Index           =   0
         Left            =   1980
         TabIndex        =   13
         Top             =   340
         Width           =   435
      End
      Begin VB.TextBox TxtOb 
         BackColor       =   &H00E8FFFF&
         Height          =   300
         Index           =   2
         Left            =   780
         MaxLength       =   10
         TabIndex        =   0
         Top             =   300
         Width           =   1035
      End
      Begin VB.CheckBox ChkActivas 
         BackColor       =   &H00FFEED9&
         Caption         =   "Ver sólo las activas"
         Height          =   195
         Left            =   4200
         TabIndex        =   3
         ToolTipText     =   "Muestra sólo las tareas no desactivadas"
         Top             =   705
         Width           =   1695
      End
      Begin VB.TextBox TxtOb 
         BackColor       =   &H00E8FFFF&
         Height          =   300
         Index           =   0
         Left            =   780
         MaxLength       =   5
         TabIndex        =   1
         Top             =   660
         Width           =   555
      End
      Begin VB.TextBox TxtOb 
         BackColor       =   &H00E8FFFF&
         Height          =   300
         Index           =   1
         Left            =   780
         MaxLength       =   150
         TabIndex        =   2
         Top             =   1020
         Width           =   5175
      End
      Begin VB.Label LblOb 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFEED9&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   345
         Width           =   450
      End
      Begin VB.Label LblOb 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFEED9&
         Caption         =   "Hora"
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   11
         Top             =   705
         Width           =   345
      End
      Begin VB.Label LblOb 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFEED9&
         Caption         =   "Tarea"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   10
         Top             =   1065
         Width           =   420
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlxDatos 
      Height          =   5235
      Left            =   60
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2040
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   9234
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   4
      FixedCols       =   0
      BackColorSel    =   14221054
      ForeColorSel    =   0
      BackColorBkg    =   16772825
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      MergeCells      =   4
      FormatString    =   $"BuscarFrm.frx":0000
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label Lbls 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click en cabecera de columna, ordena. Dobleclick para acceder al día de la tarea"
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
      Height          =   435
      Left            =   60
      TabIndex        =   8
      Top             =   1740
      Width           =   13335
   End
End
Attribute VB_Name = "BuscarFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Aqui tenemos una conexión independiente de la del otro formulario, por lo que no
'hay problemas al salir se cierra esta conexión

Dim BD As Database
Dim blnOrdenFlex As Boolean

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
  FlxDatos.FormatString = "^Fecha             |^Hora     |Tarea                                                                                                                                                                                                                                       |^Act"
  FlxDatos.Rows = 2
End Sub

'Crea sql
Private Function CreaSQL() As String
  Dim Sql As String
  
  Sql = "select FECHA, HORA, TAREA, ACTIVA" _
      & "  from TAREAS" _
      & " where 1 = 1"
  If TxtOb(2) <> "" Then
    Sql = Sql & " and FECHA "
    If OptBuscar(0).Value Then Sql = Sql & " < "
    If OptBuscar(1).Value Then Sql = Sql & " = "
    If OptBuscar(2).Value Then Sql = Sql & " > "
    Sql = Sql & "#" & Format(TxtOb(2), "mm/dd/yyyy") & "#"
  End If
  If TxtOb(0) <> "" Then Sql = Sql & " and HORA Like '*" & TxtOb(0) & "*'"
  If TxtOb(1) <> "" Then Sql = Sql & " and TAREA Like '*" & TxtOb(1) & "*'"
  If ChkActivas.Value = vbChecked Then Sql = Sql & " and ACTIVA = 'S'"
  Sql = Sql & " order by FECHA, HORA"

  CreaSQL = Sql
End Function

'Rellena el flex de datos
Private Sub RefrescaFlex()
  Dim Rec As Recordset
  Dim I As Integer
  
  Screen.MousePointer = vbHourglass
  
  Set Rec = BD.OpenRecordset(CreaSQL)
  
  InicializaFlex
  I = 0
  While Not Rec.EOF
    I = I + 1
    
    FlxDatos.TextMatrix(I, 0) = Rec!FECHA
    FlxDatos.TextMatrix(I, 1) = Rec!HORA
    FlxDatos.TextMatrix(I, 2) = Rec!TAREA
    FlxDatos.TextMatrix(I, 3) = IIf(Rec!ACTIVA = "S", "*", "")
    
    Rec.MoveNext
    
    'Si va a haber mas filas, añadir una row más
    If Not Rec.EOF Then FlxDatos.Rows = FlxDatos.Rows + 1
  Wend
  Rec.Close

  TxtOb(0).SetFocus
  FlxDatos.Row = 1
  
  Set Rec = Nothing
  Screen.MousePointer = vbDefault
End Sub
    
'Si se hace click en la cabecera, ordenar por esa columna
Private Sub FlxDatos_Click()
  With FlxDatos
    If .MouseRow < 1 Then
      blnOrdenFlex = Not blnOrdenFlex
      
      If .MouseCol = 0 Then
        .Col = .MouseCol
        .Sort = 9
      Else
        .Col = .MouseCol
        If blnOrdenFlex Then
           .Sort = flexSortStringNoCaseAscending
        Else
           .Sort = flexSortStringNoCaseDescending
        End If
      End If
      
      .Refresh
    End If
  End With
End Sub

'Método de ordenación para el campo fecha
Private Sub FlxDatos_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
  With FlxDatos
    If blnOrdenFlex Then
      Cmp = IIf(CDate(.TextMatrix(Row1, 0)) < CDate(.TextMatrix(Row2, 0)), 1, -1)
    Else
      Cmp = IIf(CDate(.TextMatrix(Row1, 0)) > CDate(.TextMatrix(Row2, 0)), 1, -1)
    End If
  End With
End Sub

'Ir al dato
Private Sub FlxDatos_DblClick()
  On Error GoTo Errores
  
  'Sólo dejar acceder a los datos si la agenda está inactiva
  If AgendaFrm.BtnActivar.Visible Then
    If FlxDatos.MouseRow > 0 Then
      If FlxDatos.TextMatrix(1, 0) <> "" Then
        AgendaFrm.Calendar.Value = FlxDatos.TextMatrix(FlxDatos.MouseRow, 0)
        Call AgendaFrm.Calendar_DateClick(FlxDatos.TextMatrix(FlxDatos.MouseRow, 0))
        Unload Me
      Else
        MsgBox "No hay datos.", vbInformation
      End If
    End If
  Else
    MsgBox "Para acceder a los datos de ese día, desactive la agenda primero.", vbInformation
  End If
  Exit Sub

Errores:
  MsgBox "Error: " & Err.Description
End Sub

Private Sub Form_Load()
  ConectaBD
  FlxDatos.MergeCol(0) = True
End Sub

'Se descarga el form, desconectar
Private Sub Form_Unload(Cancel As Integer)
  DesconectaBD
End Sub

Private Sub BtnAceptar_Click()
  RefrescaFlex
End Sub

Private Sub BtnImprimir_Click()
  Dim Rec As Recordset
  
  Screen.MousePointer = vbHourglass
  
  Set Rec = BD.OpenRecordset(CreaSQL)
  
  If Not Rec.EOF Then ImpresionTareas Rec, True
  
  Rec.Close
  Set Rec = Nothing
  Screen.MousePointer = vbDefault
End Sub

Private Sub BtnSalir_Click()
  Unload Me
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
    Case 0
      Select Case KeyAscii
        Case Asc("0") To Asc("9")
          'Se ha pulsado un nº
        Case Asc(":")
          'Se ha pulsado un separador
        Case Asc(".")
          'Se ha pulsado un separador, cambiar por :
          KeyAscii = Asc(":")
        Case 8
          'Se ha pulsado BCKSPC
        Case Else
          KeyAscii = 0
          Beep
      End Select
    
    Case 2
      Select Case KeyAscii
        Case Asc("0") To Asc("9")
          'Se ha pulsado un nº
        Case Asc(":"), Asc("."), Asc("-")
          'Se ha pulsado un separador, cambiar por /
          KeyAscii = Asc("/")
        Case Asc("/")
          'Se ha pulsado un separador
        Case 8
          'Se ha pulsado BCKSPC
        Case Else
          KeyAscii = 0
          Beep
      End Select
    
  End Select
End Sub

'Validación de un TextBox TxtOb antes de salir
Private Sub TxtOb_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index
    Case 2
      If TxtOb(Index) <> "" Then
        If Not IsDate(TxtOb(Index)) Then
          MsgBox "Esa fecha no es válida.", vbInformation
          Cancel = True
        Else
          TxtOb(Index) = Format(TxtOb(Index), "dd/mm/yyyy")
        End If
      End If
    
  End Select
End Sub
