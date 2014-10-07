Attribute VB_Name = "ImpresionMod"
Option Explicit

'Especifica los valores para el tipo de fuente de un título
Private Sub FuenteTitulo()
  Printer.FontName = "Arial"
  Printer.FontBold = True
  Printer.FontItalic = True
  Printer.FontUnderline = False
  Printer.FontSize = 16
End Sub

'Especifica los valores para el tipo de fuente de un cabecera
Private Sub FuenteCabecera()
  Printer.FontName = "Courier New"
  Printer.FontBold = True
  Printer.FontItalic = False
  Printer.FontUnderline = True
  Printer.FontSize = 8
End Sub

'Especifica los valores para el tipo de fuente normal
Private Sub FuenteNormal()
  Printer.FontName = "Courier New"
  Printer.FontBold = False
  Printer.FontItalic = False
  Printer.FontUnderline = False
  Printer.FontSize = 8
End Sub

'Proceso para imprimir un Rec pasado con datos de Tareas
Public Sub ImpresionTareas(ByRef Rec As Recordset, ByRef blnApaisado As Boolean)
  Const kMargenSup = 0.5
  Const kMargenIzq = 1
  Const kIntervalo = 0.3
  Dim PosY As Single
  Dim intNumPag As Integer
  
  If Rec.EOF Then
    MsgBox "Sin datos para imprimir.", vbInformation
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
  
  Printer.ScaleMode = vbCentimeters
  If blnApaisado Then Printer.Orientation = vbPRORLandscape
  
  'Título
  FuenteTitulo
  Printer.CurrentY = kMargenSup
  Printer.CurrentX = kMargenIzq + IIf(blnApaisado, 7 + 4, 7)
  Printer.Print "Listado de Tareas"
  
  PosY = kMargenSup + 1
  
  'Cabecera
  FuenteCabecera
  Printer.CurrentX = kMargenIzq
  Printer.CurrentY = PosY
  Printer.Print "   Fecha   -  Hora - Act - Tarea"
  
  intNumPag = 1
  FuenteNormal
  BuscarFrm.Lbls.Caption = "Imprimiendo .."
  While Not Rec.EOF
    BuscarFrm.Lbls.Caption = BuscarFrm.Lbls.Caption & "."
    DoEvents
    PosY = PosY + kIntervalo
    
    Printer.CurrentY = PosY
    Printer.CurrentX = kMargenIzq
    Printer.Print Format(Rec!FECHA, "dd/mm/yyyy") & " - " & Format(Rec!HORA, "@@@@@") & " - " & IIf(Rec!ACTIVA = "S", " * ", "   ") & " - " & Rec!TAREA
    
    If PosY > IIf(blnApaisado, 18.5, 27.5) Then
      Printer.CurrentY = IIf(blnApaisado, 19, 28)
      Printer.CurrentX = kMargenIzq + IIf(blnApaisado, 17 + 8, 17)
      Printer.Print "Página " & intNumPag
      
      Printer.NewPage
      DoEvents
      
      'Ya sin título
      PosY = kMargenSup
      intNumPag = intNumPag + 1
    
      FuenteCabecera
      Printer.CurrentX = kMargenIzq
      Printer.CurrentY = PosY
      Printer.Print "   Fecha   -  Hora - Act - Tarea"
      FuenteNormal
      BuscarFrm.Lbls.Caption = "Imprimiendo .."
    End If
    
    Rec.MoveNext
  Wend
  
  Printer.CurrentY = IIf(blnApaisado, 19, 28)
  Printer.CurrentX = kMargenIzq + IIf(blnApaisado, 17 + 8, 17)
  Printer.Print "Página " & intNumPag
      
  Printer.EndDoc
  
  BuscarFrm.Lbls.Caption = "Datos encontrados. Dobleclick para acceder al día de la tarea."
  
  'Dejar el papel como estaba
  If blnApaisado Then Printer.Orientation = vbPRORPortrait
  
  Screen.MousePointer = vbDefault
End Sub

'Proceso para imprimir un Rec pasado con datos de Tlfns
Public Sub ImpresionTlfns(ByRef Rec As Recordset, ByRef blnApaisado As Boolean)
  Const kMargenSup = 0.5
  Const kMargenIzq = 1
  Const kIntervalo = 0.3
  Dim PosY As Single
  Dim intNumPag As Integer
  
  If Rec.EOF Then
    MsgBox "Sin datos para imprimir.", vbInformation
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
  
  Printer.ScaleMode = vbCentimeters
  If blnApaisado Then Printer.Orientation = vbPRORLandscape
  
  'Título
  FuenteTitulo
  Printer.CurrentY = kMargenSup
  Printer.CurrentX = kMargenIzq + IIf(blnApaisado, 7 + 4, 7)
  Printer.Print "Listado de la Agenda"
  
  PosY = kMargenSup + 1
  
  'Cabecera
  FuenteCabecera
  Printer.CurrentX = kMargenIzq
  Printer.CurrentY = PosY
  Printer.Print "Nombre                    - Tlfn/Descripción"
  
  intNumPag = 1
  FuenteNormal
  TlfnsFrm.Lbls.Caption = "Imprimiendo .."
  While Not Rec.EOF
    TlfnsFrm.Lbls.Caption = TlfnsFrm.Lbls.Caption & "."
    DoEvents
    PosY = PosY + kIntervalo
    
    Printer.CurrentY = PosY
    Printer.CurrentX = kMargenIzq
    Printer.Print Rec!NOMBRE & Space(25 - Len(Rec!NOMBRE)) & " - " & Rec!DATOS
    
    If PosY > IIf(blnApaisado, 18.5, 27.5) Then
      Printer.CurrentY = IIf(blnApaisado, 19, 28)
      Printer.CurrentX = kMargenIzq + IIf(blnApaisado, 17 + 8, 17)
      Printer.Print "Página " & intNumPag
      
      Printer.NewPage
      DoEvents
      
      'Ya sin título
      PosY = kMargenSup
      intNumPag = intNumPag + 1
    
      FuenteCabecera
      Printer.CurrentX = kMargenIzq
      Printer.CurrentY = PosY
      Printer.Print "Nombre                    - Tlfn/Descripción"
      FuenteNormal
      TlfnsFrm.Lbls.Caption = "Imprimiendo .."
    End If
    
    Rec.MoveNext
  Wend
  
  Printer.CurrentY = IIf(blnApaisado, 19, 28)
  Printer.CurrentX = kMargenIzq + IIf(blnApaisado, 17 + 8, 17)
  Printer.Print "Página " & intNumPag
      
  Printer.EndDoc
  
  TlfnsFrm.Lbls.Caption = "Datos existentes"
  
  'Dejar el papel como estaba
  If blnApaisado Then Printer.Orientation = vbPRORPortrait
  
  Screen.MousePointer = vbDefault
End Sub
