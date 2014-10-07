VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AgendaFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SimpleAgenda 3.0"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9285
   Icon            =   "AgendaFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog DlgSonido 
      Left            =   2700
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox ChkAuto 
      BackColor       =   &H00FFEED9&
      Caption         =   "Autoactivar al inicio"
      Height          =   195
      Left            =   3000
      TabIndex        =   22
      ToolTipText     =   "Automáticamente activa la agenda al ser iniciada"
      Top             =   1140
      Width           =   1755
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2760
      Top             =   1020
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlxDatos 
      Height          =   2715
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3000
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   4789
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   4
      BackColorSel    =   14221054
      ForeColorSel    =   0
      BackColorBkg    =   16772825
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   $"AgendaFrm.frx":0442
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSComCtl2.MonthView Calendar 
      Height          =   2370
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   180
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12648447
      Appearance      =   1
      MonthBackColor  =   14221054
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   16318466
      TitleBackColor  =   15838027
      TitleForeColor  =   0
      TrailingForeColor=   12632256
      CurrentDate     =   37166
      MaxDate         =   73415
      MinDate         =   36526
   End
   Begin VB.PictureBox PicBotones 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   2940
      ScaleHeight     =   900
      ScaleWidth      =   5160
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   5160
      Begin VB.CommandButton BtnTlfns 
         BackColor       =   &H00FFEED9&
         Caption         =   "&Ver Teléfonos"
         Height          =   375
         Left            =   1770
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   1635
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00FFEED9&
         Caption         =   "&Modificar Tarea"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1635
      End
      Begin VB.CommandButton BtnBorrar 
         BackColor       =   &H00FFEED9&
         Caption         =   "&Eliminar Tarea"
         Height          =   375
         Left            =   1770
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   1635
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00FFEED9&
         Caption         =   "&Añadir Tarea"
         Height          =   375
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   1635
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00FFEED9&
         Caption         =   "&Salir"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1635
      End
      Begin VB.CommandButton BtnActivar 
         BackColor       =   &H00FFEED9&
         Caption         =   "Ac&tivar la Agenda"
         Height          =   375
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   60
         Width           =   1635
      End
      Begin VB.CommandButton BtnParar 
         BackColor       =   &H00FFEED9&
         Caption         =   "&Detener la Agenda"
         Height          =   375
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   1635
      End
   End
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
      Height          =   1575
      Left            =   3060
      TabIndex        =   18
      Top             =   1035
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox ChkRepetir 
         BackColor       =   &H00FFEED9&
         Caption         =   "Repetir"
         Height          =   195
         Left            =   1740
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton BtnSonido 
         BackColor       =   &H00FFEED9&
         Caption         =   "..."
         Height          =   300
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1000
         Width           =   315
      End
      Begin VB.TextBox TxtOb 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   780
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   8
         Text            =   "Opcional: seleccione el archivo de sonido a reproducir   --------------->"
         Top             =   1020
         Width           =   4875
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00FFEED9&
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   360
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   260
         Width           =   1455
      End
      Begin VB.CommandButton BtnAceptar 
         BackColor       =   &H00FFEED9&
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   360
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   260
         Width           =   1455
      End
      Begin VB.TextBox TxtOb 
         BackColor       =   &H00E8FFFF&
         Height          =   300
         Index           =   1
         Left            =   780
         MaxLength       =   200
         TabIndex        =   7
         Text            =   "Escriba aquí la descripción de la tarea..."
         Top             =   660
         Width           =   5175
      End
      Begin VB.TextBox TxtOb 
         BackColor       =   &H00E8FFFF&
         Height          =   300
         Index           =   0
         Left            =   780
         MaxLength       =   5
         TabIndex        =   6
         Text            =   "12:00"
         Top             =   300
         Width           =   555
      End
      Begin VB.Label LblAviso 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "No se modifican las tareas repetidas derivadas. Debe borrar las tareas derivadas después."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   1350
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Label LblOb 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFEED9&
         Caption         =   "Sonido"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   23
         Top             =   1065
         Width           =   495
      End
      Begin VB.Label LblOb 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFEED9&
         Caption         =   "Tarea"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   20
         Top             =   705
         Width           =   420
      End
      Begin VB.Label LblOb 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFEED9&
         Caption         =   "Hora"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   345
         Width           =   345
      End
   End
   Begin VB.Label LblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "http:// manuel.conde. name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F1AB4B&
      Height          =   615
      Left            =   8160
      TabIndex        =   21
      ToolTipText     =   "Pinchar para ir a la página web"
      Top             =   300
      Width           =   1155
   End
   Begin VB.Label Lbls 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Avisos para la Fecha Seleccionada. Doble click modifica / activa el aviso"
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
      Left            =   120
      TabIndex        =   17
      Top             =   2700
      Width           =   9075
   End
   Begin VB.Menu MenuPrincipal 
      Caption         =   "Menú Principal"
      Begin VB.Menu Menu 
         Caption         =   "Activar"
         Index           =   0
      End
      Begin VB.Menu Menu 
         Caption         =   "Teléfonos"
         Index           =   1
      End
      Begin VB.Menu Menu 
         Caption         =   "Salir"
         Index           =   2
      End
      Begin VB.Menu Menu 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu Menu 
         Caption         =   "Buscar"
         Index           =   4
      End
      Begin VB.Menu Menu 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu Menu 
         Caption         =   "Ayuda"
         Index           =   6
      End
   End
   Begin VB.Menu AcercaDe 
      Caption         =   "Acerca de..."
   End
End
Attribute VB_Name = "AgendaFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Está hecho con DAO porque la DLL ocupa 5.5 Mb y con ADO ocupa 8 Mb
'Además así no emplea ODBC para conectarse

Dim Sql As String
Public BD As Database

'Indica si la BD está conectada o desconectada
Dim mblnConectado As Boolean
'Modificando un dato
Dim mblnModificando As Boolean

'Guarda el path de la aplicación
Dim mstrPath As String

'Guarda la fecha en que fue lanzada la agenda, para comprobar si cambió y cambiar de día
Dim mdatFecha As Date

'Conecta con la BD
Private Function ConectaBD() As Boolean
  On Local Error Resume Next
  
  Set BD = OpenDatabase(mstrPath & "Agenda.mdb")
  
  If Err.Number <> 0 Then
    MsgBox "Se produjo un error al intentar abrir la Base de Datos: " & Err.Description
    MsgBox "Asegúrese de que el archivo 'Agenda.mdb' se encuentra en el directorio " & vbCrLf & vbCrLf & mstrPath
    ConectaBD = False
  Else
    mblnConectado = True
    ConectaBD = True
  End If
End Function

'Conecta con la BD
Private Sub DesconectaBD()
  'Liquidar la conexión para que no se empleen recursos del sistema
  On Local Error Resume Next
  mblnConectado = False
  BD.Close
  Set BD = Nothing
End Sub

'Comprueba si la fecha está en formato de 4 dígitos
Public Function FormatoAnhoOk() As Boolean
  Dim strFecha As String
  Dim intLongitud As Integer

  strFecha = Date
  FormatoAnhoOk = IIf(IsNumeric(Right(strFecha, 4)), True, False)
End Function

'Configura los botones y opciones activas
Private Sub ConfiguraEstado(blnActivo As Boolean)
  BtnActivar.Visible = Not blnActivo
  BtnParar.Visible = blnActivo
  
  Calendar.Enabled = Not blnActivo
  FlxDatos.Enabled = Not blnActivo
  'BtnAñadir.Enabled = Not blnActivo
  BtnBorrar.Enabled = Not blnActivo
  BtnModificar.Enabled = Not blnActivo
  
  Menu(0).Caption = IIf(blnActivo, "Desactivar", "Activar")
End Sub

'Comprueba que una elemento en un flex pasado no haya sido seleccionado ya previamente
Private Function ExisteElementoEnFlex(FlxGrid As MSHFlexGrid, strValor As String) As Boolean
  Dim I As Integer

  ExisteElementoEnFlex = False

  For I = 1 To FlxGrid.Rows - 1
    If FlxGrid.TextMatrix(I, 0) = strValor Then
      ExisteElementoEnFlex = True
      Exit Function
    End If
  Next I
End Function

'Resetea el flex de datos
Private Sub InicializaFlex()
  FlxDatos.Clear
  FlxDatos.FormatString = "^Hora     |Tarea                                                                                        |Sonido                                                      |^Activa"
  FlxDatos.Rows = 2
End Sub

'Rellena el flex de datos
Private Sub RefrescaFlex()
  Dim Rec As Recordset
  Dim I As Integer
  
  Screen.MousePointer = vbHourglass
  
  Sql = "select FECHA, HORA, TAREA, SONIDO, ACTIVA" _
      & "  from TAREAS" _
      & " where FECHA = DateValue('" & Calendar.Value & "')" _
      & " order by HORA"
  Set Rec = BD.OpenRecordset(Sql)
  
  InicializaFlex
  I = 0
  While Not Rec.EOF
    I = I + 1
    
    FlxDatos.TextMatrix(I, 0) = Rec!HORA
    FlxDatos.TextMatrix(I, 1) = Rec!TAREA
    FlxDatos.TextMatrix(I, 2) = "" & Rec!SONIDO
    FlxDatos.TextMatrix(I, 3) = IIf(Rec!ACTIVA = "S", "*", "")
    
    Rec.MoveNext
    
    'Si va a haber mas filas, añadir una row más
    If Not Rec.EOF Then FlxDatos.Rows = FlxDatos.Rows + 1
  Wend
  Rec.Close

  If FlxDatos.Visible And FlxDatos.Enabled Then FlxDatos.SetFocus
  'FlxDatos.Row = 1
  
  Set Rec = Nothing
  Screen.MousePointer = vbDefault
End Sub

'Cambiar muestra
Private Sub FlxDatos_DblClick()
  On Error GoTo Errores
  
  If FlxDatos.TextMatrix(1, 0) <> "" Then
    If FlxDatos.TextMatrix(0, FlxDatos.MouseCol) = "Activa" Then
      If FlxDatos.TextMatrix(FlxDatos.MouseRow, 3) = "*" Then
        'Actualizar el flex
        FlxDatos.TextMatrix(FlxDatos.MouseRow, 3) = ""
        
        Sql = "update TAREAS" _
            & "   set ACTIVA = 'N'" _
            & " where FECHA = DateValue('" & Calendar.Value & "')" _
            & "   and HORA = '" & FlxDatos.TextMatrix(FlxDatos.MouseRow, 0) & "'"
        BD.Execute Sql
      Else
        'Actualizar el flex
        FlxDatos.TextMatrix(FlxDatos.MouseRow, 3) = "*"
        
        Sql = "update TAREAS" _
            & "   set ACTIVA = 'S'" _
            & " where FECHA = DateValue('" & Calendar.Value & "')" _
            & "   and HORA = '" & FlxDatos.TextMatrix(FlxDatos.MouseRow, 0) & "'"
        BD.Execute Sql
      End If
    
    Else
      'Modificar la tarea
      If BtnModificar.Enabled Then Call BtnModificar_Click
    End If
  Else
    MsgBox "No hay datos.", vbInformation
  End If
  Exit Sub

Errores:
  MsgBox "Error: " & Err.Description
End Sub

'Cuando se elige una fecha, buscar los datos para mostrarlos
Public Sub Calendar_DateClick(ByVal DateClicked As Date)
  RefrescaFlex
End Sub

'Acciones cuando se lanza el timer
Private Sub Timer_Timer()
  Dim I As Integer
  Dim strCantidad As String
  Dim strHoras As String, strMinutos As String
  Dim blnRet As Boolean, wFlags As Integer
  
  'Si no cambió el día seguir normalmente, sino cambiar el día de la agenda
  'y refrescar tareas
  If Date <> mdatFecha Then
    ConectaBD
    
    Calendar = Date
    RefrescaFlex
    
    DesconectaBD
    
    mdatFecha = Date
  
    'Actualizar la lista de tareas
    If FlxDatos.TextMatrix(1, 0) = "" Then
      nId.szTip = "Agenda comprobando 0 tareas" & vbNullChar
    Else
      nId.szTip = "Agenda comprobando " & FlxDatos.Rows - 1 & " tareas entre las " & FlxDatos.TextMatrix(1, 0) & " y las " & FlxDatos.TextMatrix(FlxDatos.Rows - 1, 0) & vbNullChar
    End If
    Shell_NotifyIcon NIM_MODIFY, nId
  End If
  
  'Comprobar todas las tareas del flex
  For I = 1 To FlxDatos.Rows - 1
    'Si la tarea está activa, comprobar si ha llegado su momento
    If FlxDatos.TextMatrix(I, 3) = "*" Then
      If FlxDatos.TextMatrix(I, 0) <= Format(Time, "HH:mm") Then
        'Es la hora exacta o mayor, avisar
        'Ejecutar el sonido, si está configurado
        If FlxDatos.TextMatrix(I, 2) <> "" Then
          wFlags = SND_ASYNC 'Or SND_NODEFAULT
          blnRet = sndPlaySound(FlxDatos.TextMatrix(I, 2), wFlags)
'          If blnRet Then
'            AvisoFrm.mEintTiempo = 10000
'            AvisoFrm.mEstrMsg = "Error con el sonido " & FlxDatos.TextMatrix(I, 2) & ": " & Err.Description
'            AvisoFrm.Show vbModal
'          End If
        End If
        
        'Mostrar aviso
        'Si se están visualizando los tlfns, descargar para que no de error
        If TlfnsFrm.Visible Then Unload TlfnsFrm
        Me.WindowState = vbNormal
        Me.Show
        AvisoFrm.mEintTiempo = 10000
        AvisoFrm.mEstrMsg = FlxDatos.TextMatrix(I, 0) & ": es la hora de " & FlxDatos.TextMatrix(I, 1)
        AvisoFrm.Show vbModal
        
        ConectaBD
        
        Sql = "update TAREAS" _
            & "   set ACTIVA = 'N'" _
            & " where FECHA = DateValue('" & Calendar.Value & "')" _
            & "   and HORA = '" & FlxDatos.TextMatrix(I, 0) & "'"
        BD.Execute Sql
          
        RefrescaFlex

Desconectar:
        'Ahorrar recursos
        DesconectaBD
      
      End If  'If FlxDatos.TextMatrix(I, 0) <= Format(Time, "HH:mm") Then
    End If  'If FlxDatos.TextMatrix(I, 2) = "*" Then
  Next I
End Sub

'Conectarse a la BD, cargar en el tray y leer config
Private Sub Form_Load()
  Dim Rec As Recordset
  Dim strCadena As String, strLogo As String
  
  'Antes de nada, verificar la fecha
  If Not FormatoAnhoOk Then
    MsgBox "Para el correcto funcionamiento de la Agenda debe configurar el formato de fecha de su sistema como dd/mm/aaaa (el año con 4 dígitos)." & vbCrLf & vbCrLf & "Vaya al 'Panel de Control -> Configuración Regional' y haga allí los cambios necesarios.", vbInformation
    End
  End If
  
  'Obtener la ruta  de trabajo
  mstrPath = App.Path
  'Si se ejecuta en local no devuelve la última barra, pero en remoto si, comprobarlo
  If Right(mstrPath, 1) <> "\" Then mstrPath = mstrPath & "\"
  
  If Not ConectaBD Then End
  
  'Leer la última configuración
  ChkAuto = IIf(GetSetting(App.EXEName, "Config", "Auto", "No") = "Si", vbChecked, vbUnchecked)
  
  'Añadir icono al tray
  nId.cbSize = Len(nId)
  nId.hwnd = Me.hwnd
  nId.uId = vbNull
  nId.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  nId.uCallBackMsg = WM_MOUSEMOVE
  nId.hIcon = Me.Icon
  nId.szTip = "Agenda parada" & vbNullChar
  
  Shell_NotifyIcon NIM_ADD, nId
  
  Calendar = Date
  RefrescaFlex
  
  'Verificar tareas anteriores sin hacer
  Sql = "select count(*) from TAREAS" _
      & " where FECHA < #" & Format(Date, "mm/dd/yyyy") & "#" _
      & "   and ACTIVA = 'S'"
  Set Rec = BD.OpenRecordset(Sql)
  
  If Not Rec.EOF Then
    If Rec(0) > 0 Then
      If MsgBox("Atención: existen " & Rec(0) & " tareas activas anteriores a hoy. ¿Desea desactivarlas (las tareas de días pasados no son comprobadas por la agenda)?", vbYesNo + vbDefaultButton2) = vbYes Then
        Sql = "update TAREAS" _
            & "   set ACTIVA = 'N'" _
            & " where FECHA < #" & Format(Date, "mm/dd/yyyy") & "#" _
            & "   and ACTIVA = 'S'"
        BD.Execute Sql
      End If
    End If
  End If
  Rec.Close
  
  'Inizializar la matriz global de Sqls
  ReDim GarrSql(0)
  
  Set Rec = Nothing
  
  If ChkAuto.Value = vbChecked Then
    Me.WindowState = vbMinimized
    BtnActivar = True
  End If
End Sub

'Foco en los botones
Private Sub Form_Activate()
  If PicBotones.Enabled And PicBotones.Visible Then PicBotones.SetFocus
End Sub

'Se pulsa una tecla
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Local Error Resume Next
  
  If KeyCode = vbKeyEscape Then
    'Minimizar la aplicación
    Me.WindowState = vbMinimized
  End If
End Sub

'Capturar el ratón sobre el tray
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Msg As Long
  
  'Al salir de la label, restablecer su color
  If LblInfo.FontUnderline Then
    LblInfo.FontUnderline = False
    LblInfo.ForeColor = &HF1AB4B
  End If
  
  If Me.ScaleMode = vbPixels Then
    Msg = X
  Else
    Msg = X / Screen.TwipsPerPixelX
  End If
  
  If Not TlfnsFrm.Visible Then
    Select Case Msg
      Case WM_LBUTTONUP
        Me.WindowState = vbNormal
        Me.Show
      Case WM_RBUTTONUP
        'Mostrar el menú de acciones, se dispara también clickando el ratón sobre el form
        'normal, evitarlo
        If Me.WindowState = vbMinimized Then PopupMenu MenuPrincipal
    End Select
  End If
End Sub

'Se minimiza o maximiza el form
Private Sub Form_Resize()
  'Al minimizar ocultar para que no aparezca minimizado sobre Inicio
  If Me.WindowState = vbMinimized Then Me.Hide
End Sub

'Se descarga el form, quitar del tray
Private Sub Form_Unload(Cancel As Integer)
  nId.cbSize = Len(nId)
  nId.hwnd = Me.hwnd
  nId.uId = vbNull
  
  Shell_NotifyIcon NIM_DELETE, nId
  
  'Guardar los últimos parámetros de uso
  SaveSetting App.EXEName, "Config", "Auto", IIf(ChkAuto.Value = vbChecked, "Si", "No")
  
  If mblnConectado Then DesconectaBD
  End
End Sub

'Opciones del menú
Private Sub Menu_Click(Index As Integer)
  Select Case Index
    Case 0
      If Menu(Index).Caption = "Activar" Then
         Call BtnActivar_Click
      Else
         Call BtnParar_Click
      End If
    
    Case 1
      Call BtnTlfns_Click
    
    Case 2
      Call BtnSalir_Click
    
    Case 4
      BuscarFrm.Show vbModal
    
    Case 6
      AyudaFrm.Show vbModal
  End Select
End Sub

'Menú AcercaDe
Private Sub AcercaDe_Click()
  AcercaDeFrm.Show vbModal
End Sub

'Iniciar el proceso
Private Sub BtnActivar_Click()
  'Si no estamos en la fecha actual, ir a ella
  If Calendar.Value <> Date Then
    Calendar = Date
    RefrescaFlex
  End If
    
  If FlxDatos.TextMatrix(1, 0) = "" And ChkAuto.Value = vbUnchecked Then
    If MsgBox("No hay ninguna tarea para hoy, no es necesario iniciar la Agenda. ¿Desea iniciarla igualmente?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
  End If
  
  'Todo Ok, comenzar cuenta
  ConfiguraEstado True
  
  'Guardar la fecha de lanzamiento
  mdatFecha = Date
  
  DesconectaBD
  
  Timer.Enabled = True

  'Minimizar la aplicación
  Me.WindowState = vbMinimized

  If FlxDatos.TextMatrix(1, 0) = "" Then
    nId.szTip = "Agenda comprobando 0 tareas" & vbNullChar
  Else
    nId.szTip = "Agenda comprobando " & FlxDatos.Rows - 1 & " tareas entre las " & FlxDatos.TextMatrix(1, 0) & " y las " & FlxDatos.TextMatrix(FlxDatos.Rows - 1, 0) & vbNullChar
  End If
  Shell_NotifyIcon NIM_MODIFY, nId
End Sub

'Para el proceso de comprobación de la agenda
Private Sub BtnParar_Click()
  ConfiguraEstado False
  
  'Intentar restablecer la conexión con la BD
  ConectaBD
  
  Timer.Enabled = False

  nId.szTip = "Agenda detenida" & vbNullChar
  Shell_NotifyIcon NIM_MODIFY, nId
End Sub

'Ver los tlfns
Private Sub BtnTlfns_Click()
  TlfnsFrm.Show vbModal
End Sub

'Terminar la aplicación
Private Sub BtnSalir_Click()
  Unload Me
End Sub

'Añadir una tarea
Private Sub BtnAñadir_Click()
  ChkAuto.Visible = False
  FraDatos.Visible = True
  mblnModificando = False
  
  'Sólo si la agenda está desactivada hacer lo siguiente
  If mblnConectado Then
    Calendar.Enabled = False
    PicBotones.Enabled = False
    FlxDatos.Enabled = False
  End If
End Sub

'Borrar tarea seleccionada
Private Sub BtnBorrar_Click()
  On Error GoTo Errores
  
  If FlxDatos.TextMatrix(1, 0) <> "" Then
    If MsgBox("¿Seguro que desea eliminar esa tarea?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then Exit Sub
    
    If MsgBox("¿Desea eliminar también las repeticiones de esta tarea, si existieran?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
      Sql = "delete from TAREAS" _
          & " where HORA = '" & FlxDatos.TextMatrix(FlxDatos.Row, 0) & "'" _
          & "   and TAREA = '" & FlxDatos.TextMatrix(FlxDatos.Row, 1) & "'" _
          & "   and REPETIDA = 'S'"
      BD.Execute Sql
    End If
    
    Sql = "delete from TAREAS" _
        & " where FECHA = DateValue('" & Calendar.Value & "')" _
        & "   and HORA = '" & FlxDatos.TextMatrix(FlxDatos.Row, 0) & "'"
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
    TxtOb(2) = FlxDatos.TextMatrix(FlxDatos.Row, 2)
    If TxtOb(2) = "" Then TxtOb(2) = "Opcional: seleccione el archivo de sonido a reproducir   --------------->"
    ChkAuto.Visible = False
    FraDatos.Visible = True
    mblnModificando = True
    Calendar.Enabled = False
    PicBotones.Enabled = False
    FlxDatos.Enabled = False
    LblAviso.Visible = True
  End If
End Sub

'Establecer el archivo de sonido
Private Sub BtnSonido_Click()
  DlgSonido.Filter = "Wave-files (*.wav)|*.wav"
  DlgSonido.DialogTitle = "Seleccione el archivo de sonido WAV ..."
  DlgSonido.ShowOpen
  TxtOb(2) = ""
  TxtOb(2) = DlgSonido.FileName
End Sub

'Añadir la tarea
Private Sub BtnAceptar_Click()
  Dim strSonido As String
  Dim I As Integer
  
  On Error GoTo Errores
  
  If TxtOb(0) <> "" And TxtOb(1) <> "" Then
    If Not mblnModificando And ExisteElementoEnFlex(FlxDatos, TxtOb(0)) Then
      MsgBox "Ya existe una tarea para la hora seleccionada, modifíquela en lugar de añadirla.", vbInformation
      Exit Sub
    End If
    
    If TxtOb(2) = "Opcional: seleccione el archivo de sonido a reproducir   --------------->" Then
      strSonido = ""
    Else
      strSonido = TxtOb(2)
    End If
    
    'Enc caso de que quiera repetir el evento, mostrar el formulario que preparará las Sqls
    If ChkRepetir.Value = vbChecked Then
      AgendaRepetirFrm.mEstrFechaInicial = Calendar.Value
      AgendaRepetirFrm.mEstrHora = TxtOb(0)
      AgendaRepetirFrm.mEstrTarea = TxtOb(1)
      AgendaRepetirFrm.mEstrSonido = strSonido
      AgendaRepetirFrm.Show vbModal
    End If
    
    If mblnModificando Then
      LblAviso.Visible = False
      
      Sql = "update TAREAS" _
          & "   set HORA = '" & TxtOb(0) & "'" _
          & "     , TAREA = '" & TxtOb(1) & "'" _
          & "     , SONIDO = '" & strSonido & "'" _
          & " where FECHA = DateValue('" & Calendar.Value & "')" _
          & "   and HORA = '" & FlxDatos.TextMatrix(FlxDatos.Row, 0) & "'"
    Else
      Sql = "insert into TAREAS" _
          & "       (FECHA, HORA, TAREA, SONIDO, ACTIVA) " _
          & "values ('" & Calendar.Value & "','" & TxtOb(0) & "','" & TxtOb(1) & "','" & strSonido & "','S')"
    End If
    
    'Si no se está conectado es que la agenda está en ejecución: conectarse (y desconectarse luego, no modificar el estado)
    If Not mblnConectado Then ConectaBD
      
    'Proceso normal, ya estamos conectados
    BD.Execute Sql
      
    'Si hay repeticiones, meterlas en la BD (sin errores, por si hay duplicidad)
    If GarrSql(0) <> "" Then
      On Error Resume Next
      For I = LBound(GarrSql) To UBound(GarrSql)
        Sql = GarrSql(I)
        BD.Execute Sql
        If Err.Number <> 0 Then MsgBox "Datos duplicados, puede que el día y hora especificados ya estén en uso. Observe la consulta y revise la agenda para el día y hora especificados: " & Sql
      Next I
    
      ReDim GarrSql(0)
      GarrSql(0) = ""
      
      On Error GoTo Errores
    End If
    
    RefrescaFlex
    
    If mblnConectado Then
      'Proceso normal, ya estamos conectados y viendo la agenda
      Calendar.Enabled = True
      PicBotones.Enabled = True
      FlxDatos.Enabled = True
      ChkAuto.Visible = True
      FraDatos.Visible = False
    Else
      'La agenda está en ejecución y ya nos conectamos antes para grabar: desconectar
      ChkAuto.Visible = True
      FraDatos.Visible = False
      DesconectaBD
    End If
    
  Else
    MsgBox "Rellene todos los campos.", vbInformation
  End If
  Exit Sub

Errores:
  MsgBox "Error: " & Err.Description
End Sub

'Cancelar añadir/modificar
Private Sub BtnCancelar_Click()
  ChkAuto.Visible = True
  FraDatos.Visible = False
  LblAviso.Visible = False
  'Sólo si la agenda está desactivada hacer lo siguiente
  If mblnConectado Then
    Calendar.Enabled = True
    FlxDatos.Enabled = True
    PicBotones.Enabled = True
    PicBotones.SetFocus
  End If
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
        Case Asc(":")
          'Se ha pulsado un separador
        Case Asc(".")
          'Se ha pulsado un separador, cambiar por :
          KeyAscii = Asc(":")
        Case vbKeyBack
          'Se ha pulsado BCKSPC
        Case Else
          KeyAscii = 0
          Beep
      End Select
    
    Case 1
      'Valen todas menos el apóstrofe
      Select Case KeyAscii
        Case Asc("'")
          'Se ha pulsado un separador
          KeyAscii = 0
          MsgBox "El apóstrofe (') no es un caracter válido. En su lugar puede emplear un acento (seguido de un espacio) si lo desea.", vbInformation
      End Select
    
    Case 2
      'Valen todas menos el apóstrofe
      Select Case KeyAscii
        Case vbKeyBack
          'Se ha pulsado una tecla de borrar, vaciar el texto
          TxtOb(Index) = ""
      End Select
    
  End Select
End Sub

Private Sub TxtOb_Validate(Index As Integer, Cancel As Boolean)
  Dim strHoras As String, strMinutos As String
  Const MsgErr = "Formato horario incorrecto (HH:MM)"
  
  Select Case Index
    Case 0
      'Comprobar el formato de hora
      If TxtOb(Index) <> "" Then
'        On Error Resume Next
        
        If InStr(1, TxtOb(Index), ":") = 0 And Len(TxtOb(Index)) > 2 Then
          MsgBox MsgErr, vbInformation
          Cancel = True
          Exit Sub
        End If
        
        'Dependiendo de lo que haya escrito, añadir o no espacios
        TxtOb(Index) = TxtOb(Index) + Space(5 - Len(TxtOb(Index)))
        
        strHoras = Left(TxtOb(Index), 2)
        strMinutos = Right(TxtOb(Index), 2)
        
        If strHoras = "" Then
          MsgBox MsgErr, vbInformation
          Cancel = True
          Exit Sub
        End If
        
        If Not IsNumeric(strHoras) Then
          MsgBox MsgErr, vbInformation
          TxtOb(Index) = RTrim(TxtOb(Index))
          Cancel = True
          Exit Sub
        ElseIf Val(strHoras) > 23 Then
          MsgBox "La hora debe estar entre 0 y 23", vbInformation
          Cancel = True
          Exit Sub
        End If
        
        strHoras = Format(strHoras, "00")
        
        'Si no ha escrito minutos, añadirlos
        If Not IsNumeric(strMinutos) Then strMinutos = "00"
        
        If Val(strMinutos) > 59 Then
          MsgBox "Los minutos deben estar entre 0 y 59", vbInformation
          Cancel = True
          Exit Sub
        End If
        
        strMinutos = Format(strMinutos, "00")
        
        TxtOb(Index) = strHoras & ":" & strMinutos
      End If
      
  End Select
End Sub

'Lanzar la web
Private Sub LblInfo_Click()
  ShellExecute hwnd, "open", "http://manuel.conde.name", vbNullString, vbNullString, SW_SHOW
End Sub

'Destacar el enlace
Private Sub LblInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not LblInfo.FontUnderline Then
    LblInfo.FontUnderline = True
    LblInfo.ForeColor = &H0
  End If
End Sub
