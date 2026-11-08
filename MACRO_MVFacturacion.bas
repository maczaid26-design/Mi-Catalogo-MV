Attribute VB_Name = "MVFacturacion"
Option Explicit

' ============================================================
'  MV BELLEZA Y ESTILO - Facturacion e inventario automatico
'  Macros: NuevaFactura  y  GuardarYRegistrarVenta
' ============================================================

Private Const HOJA_PLANTILLA As String = "PLANTILLA FACTURA"
Private Const HOJA_VENTAS As String = "VENTAS"
Private Const HOJA_PRODUCTOS As String = "PRODUCTOS"
Private Const HOJA_INDICE As String = "FACTURAS"
Private Const PRIMERA_LINEA As Long = 9
Private Const ULTIMA_FILA_VENTAS As Long = 5000
Private Const ULTIMA_FILA_PRODUCTOS As Long = 569
Private Const TITULO As String = "MV BELLEZA Y ESTILO"

' ------------------------------------------------------------
'  Utilidades
' ------------------------------------------------------------
Private Function TextoCelda(ByVal rng As Range) As String
    ' Devuelve el texto de la celda sin fallar si tiene error
    On Error Resume Next
    TextoCelda = Trim(CStr(rng.Value))
    If Err.Number <> 0 Then TextoCelda = ""
    On Error GoTo 0
End Function

Private Function EsHojaFactura(ByVal ws As Worksheet) As Boolean
    ' Hoja cuyo nombre empieza por "FACTURA " (no la plantilla)
    EsHojaFactura = (UCase(Left(ws.Name, 8)) = "FACTURA ") And (UCase(ws.Name) <> UCase(HOJA_PLANTILLA))
End Function

Private Function EsFormatoNuevo(ByVal ws As Worksheet) As Boolean
    ' Formato nuevo = tiene la columna CODIGO en B8
    EsFormatoNuevo = (UCase(TextoCelda(ws.Range("B8"))) Like "C*DIGO")
End Function

Private Function NumeroValido(ByVal num As String) As Boolean
    NumeroValido = (num Like "MV-###*") And (num <> "MV-000")
End Function

Private Function NumeroDeFactura(ByVal ws As Worksheet) As Long
    Dim t As String
    t = UCase(TextoCelda(ws.Range("B5")))
    t = Replace(t, "MV-", "")
    t = Replace(t, "MV", "")
    NumeroDeFactura = CLng(Val(t))
End Function

Private Function ExisteNumeroEnOtraHoja(ByVal ws As Worksheet, ByVal num As String) As Boolean
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> ws.Name Then
            If EsHojaFactura(sh) Then
                If UCase(TextoCelda(sh.Range("B5"))) = num Then
                    ExisteNumeroEnOtraHoja = True
                    Exit Function
                End If
            End If
        End If
    Next sh
End Function

Private Function NombreHojaSeguro(ByVal n As Long, ByVal cliente As String) As String
    Dim t As String, base As String, k As Long, c As String
    t = ""
    For k = 1 To Len(cliente)
        c = Mid(cliente, k, 1)
        If InStr("\/?*[]:'", c) = 0 Then t = t & c
    Next k
    t = Trim(UCase(t))
    base = "FACTURA " & CStr(n)
    If t <> "" Then base = base & " " & t
    If Len(base) > 31 Then base = Trim(Left(base, 31))
    NombreHojaSeguro = base
End Function

Private Function ExisteHoja(ByVal nombre As String) As Boolean
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If UCase(sh.Name) = UCase(nombre) Then
            ExisteHoja = True
            Exit Function
        End If
    Next sh
End Function

Private Function SiguienteNumero() As Long
    Dim sh As Worksheet, mx As Long, v As Long
    mx = 0
    For Each sh In ThisWorkbook.Worksheets
        If EsHojaFactura(sh) Then
            v = NumeroDeFactura(sh)
            If v > mx Then mx = v
        End If
    Next sh
    SiguienteNumero = mx + 1
End Function

' ------------------------------------------------------------
'  Indice (hoja FACTURAS)
' ------------------------------------------------------------
Private Sub AgregarFilaIndice(ByVal wsF As Worksheet)
    Dim wsI As Worksheet, filaTot As Long, f As Long, k As Long
    Dim q As String, nom As String
    Set wsI = ThisWorkbook.Worksheets(HOJA_INDICE)
    nom = wsF.Name
    For k = 5 To 600
        If UCase(Trim(CStr(wsI.Cells(k, 9).Value))) = UCase(nom) Then Exit Sub
    Next k
    filaTot = 0
    For k = 5 To 600
        If UCase(Trim(CStr(wsI.Cells(k, 1).Value))) = "TOTALES" Then
            filaTot = k
            Exit For
        End If
    Next k
    If filaTot = 0 Then Exit Sub
    wsI.Rows(filaTot).Insert Shift:=xlDown
    f = filaTot
    q = "'" & Replace(nom, "'", "''") & "'!"
    wsI.Cells(f, 1).Formula = "=" & q & "B5"
    wsI.Cells(f, 2).Formula = "=" & q & "B6"
    wsI.Cells(f, 3).Formula = "=" & q & "E5"
    wsI.Cells(f, 4).Formula = "=" & q & "F20"
    wsI.Cells(f, 5).Formula = "=" & q & "D30"
    wsI.Cells(f, 6).Formula = "=" & q & "D31"
    wsI.Cells(f, 7).Formula = "=" & q & "D32"
    wsI.Cells(f, 8).Formula = "=IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(A" & f & ",""MV-"",""""),""MV"","""")),0)"
    wsI.Cells(f, 9).Value = nom
    wsI.Cells(f, 10).Formula = "=" & q & "$H$10"
    wsI.Cells(f, 11).Value = 1
    wsI.Cells(f + 1, 4).Formula = "=SUM(D5:D" & f & ")"
    wsI.Cells(f + 1, 5).Formula = "=SUM(E5:E" & f & ")"
    wsI.Cells(f + 1, 6).Formula = "=SUM(F5:F" & f & ")"
End Sub

' ------------------------------------------------------------
'  BOTON: NUEVA FACTURA
' ------------------------------------------------------------
Public Sub NuevaFactura()
    Dim wsPl As Worksheet, wsN As Worksheet, sh As Worksheet, wsDespues As Worksheet
    Dim n As Long, num As String, nombre As String, resp As Variant, cliente As String
    Dim calc As Long, k As Long

    On Error GoTo Falla
    If Not ExisteHoja(HOJA_PLANTILLA) Then
        MsgBox "No encuentro la hoja " & HOJA_PLANTILLA & ". No la borres ni le cambies el nombre.", vbExclamation, TITULO
        Exit Sub
    End If
    Set wsPl = ThisWorkbook.Worksheets(HOJA_PLANTILLA)

    n = SiguienteNumero()
    num = "MV-" & Format(n, "000")

    resp = Application.InputBox(Prompt:="Nombre del cliente para la factura " & num & ":", Title:="Nueva factura " & num, Type:=2)
    If VarType(resp) = vbBoolean Then Exit Sub
    cliente = Trim(CStr(resp))

    nombre = NombreHojaSeguro(n, cliente)
    k = 2
    Do While ExisteHoja(nombre)
        nombre = Left(NombreHojaSeguro(n, cliente), 27) & " (" & CStr(k) & ")"
        k = k + 1
    Loop

    ' la nueva hoja va despues de la ultima factura
    Set wsDespues = ThisWorkbook.Worksheets(HOJA_INDICE)
    For Each sh In ThisWorkbook.Worksheets
        If EsHojaFactura(sh) Then Set wsDespues = sh
    Next sh

    Application.ScreenUpdating = False
    calc = Application.Calculation
    Application.Calculation = xlCalculationManual

    wsPl.Copy After:=wsDespues
    Set wsN = ActiveSheet
    wsN.Name = nombre
    wsN.Range("B5").Value = num
    wsN.Range("E5").Value = Date
    wsN.Range("E5").NumberFormat = "dd/mm/yyyy"
    wsN.Range("B6").Value = cliente
    wsN.Range("E6").ClearContents
    wsN.Range("B9:B18").ClearContents
    wsN.Range("D9:D18").ClearContents
    wsN.Range("B24:E28").ClearContents

    AgregarFilaIndice wsN

    Application.Calculation = calc
    Application.Calculate
    Application.ScreenUpdating = True
    wsN.Activate
    wsN.Range("B9").Select
    MsgBox "Factura " & num & " creada." & vbCrLf & _
           "Escribe los códigos y cantidades, y luego presiona GUARDAR Y REGISTRAR VENTA.", vbInformation, TITULO
    Exit Sub

Falla:
    Application.ScreenUpdating = True
    If calc <> 0 Then
        On Error Resume Next
        Application.Calculation = calc
        On Error GoTo 0
    End If
    MsgBox "No se pudo crear la factura: " & Err.Description, vbCritical, TITULO
End Sub

' ------------------------------------------------------------
'  BOTON: GUARDAR Y REGISTRAR VENTA
' ------------------------------------------------------------
Public Sub GuardarYRegistrarVenta()
    Dim ws As Worksheet, wsV As Worksheet, wsP As Worksheet
    Dim num As String, medio As String, errores As String, avisos As String
    Dim fecha As Date
    Dim i As Long, j As Long, r As Long, idx As Long, nFilas As Long
    Dim nLineas As Long, nNuevas As Long, nAct As Long, nBorr As Long, unidades As Double
    Dim cod(1 To 10) As String
    Dim cant(1 To 10) As Double
    Dim pre(1 To 10) As Double
    Dim lista(1 To 10) As Double
    Dim filaP(1 To 10) As Long
    Dim usa(1 To 10) As Boolean
    Dim filaDest(1 To 10) As Long
    Dim vK As Variant, vB As Variant, vD As Variant
    Dim borrar() As Boolean
    Dim ocupada() As Boolean
    Dim mia As Boolean
    Dim llave As String, previa As Double, necesita As Double, hay As Double, k As Long
    Dim pos As Variant, calc As Long
    Dim hechoCalc As Boolean

    On Error GoTo Falla
    Set ws = ActiveSheet
    If Not EsHojaFactura(ws) Then
        MsgBox "Este botón solo funciona dentro de una factura (hoja 'FACTURA n ...').", vbExclamation, TITULO
        Exit Sub
    End If
    If Not EsFormatoNuevo(ws) Then
        MsgBox "Esta factura es histórica (formato anterior, sin columna CÓDIGO) y no se registra con el botón." & vbCrLf & _
               "Crea las nuevas con el botón NUEVA FACTURA.", vbInformation, TITULO
        Exit Sub
    End If

    num = UCase(TextoCelda(ws.Range("B5")))
    If Not NumeroValido(num) Then
        MsgBox "El número de factura (celda B5) debe tener la forma MV-031 y no puede ser MV-000.", vbExclamation, TITULO
        Exit Sub
    End If
    If Not IsDate(ws.Range("E5").Value) Then
        MsgBox "Escribe la fecha de la factura en la celda E5.", vbExclamation, TITULO
        Exit Sub
    End If
    fecha = CDate(ws.Range("E5").Value)
    If ExisteNumeroEnOtraHoja(ws, num) Then
        MsgBox "El número " & num & " ya existe en otra hoja de factura." & vbCrLf & _
               "Cámbialo en la celda B5 para no mezclar ventas.", vbExclamation, TITULO
        Exit Sub
    End If

    Set wsV = ThisWorkbook.Worksheets(HOJA_VENTAS)
    Set wsP = ThisWorkbook.Worksheets(HOJA_PRODUCTOS)

    ' ---------- 1. leer y validar las lineas de la factura
    nLineas = 0
    errores = ""
    For i = 1 To 10
        r = PRIMERA_LINEA + i - 1
        cod(i) = TextoCelda(ws.Cells(r, 2))
        usa(i) = False
        If cod(i) <> "" Then
            pos = Application.Match(cod(i), wsP.Range("A2:A" & ULTIMA_FILA_PRODUCTOS), 0)
            If IsError(pos) Then
                errores = errores & "Línea " & i & ": el código '" & cod(i) & "' no existe en PRODUCTOS." & vbCrLf
            Else
                filaP(i) = CLng(pos) + 1
                If Not IsNumeric(ws.Cells(r, 4).Value) Or IsEmpty(ws.Cells(r, 4).Value) Or CStr(ws.Cells(r, 4).Value) = "" Then
                    errores = errores & "Línea " & i & ": falta la cantidad." & vbCrLf
                ElseIf CDbl(ws.Cells(r, 4).Value) <= 0 Or CDbl(ws.Cells(r, 4).Value) <> Int(CDbl(ws.Cells(r, 4).Value)) Then
                    errores = errores & "Línea " & i & ": la cantidad debe ser un número entero mayor que 0." & vbCrLf
                ElseIf Not IsNumeric(ws.Cells(r, 5).Value) Or CStr(ws.Cells(r, 5).Value) = "" Then
                    errores = errores & "Línea " & i & ": el precio no es válido." & vbCrLf
                Else
                    cant(i) = CDbl(ws.Cells(r, 4).Value)
                    pre(i) = CDbl(ws.Cells(r, 5).Value)
                    lista(i) = CDbl(wsP.Cells(filaP(i), 5).Value)
                    usa(i) = True
                    nLineas = nLineas + 1
                End If
            End If
        ElseIf TextoCelda(ws.Cells(r, 4)) <> "" Then
            errores = errores & "Línea " & i & ": hay una cantidad pero falta el código." & vbCrLf
        End If
    Next i
    If errores <> "" Then
        MsgBox "No se registró nada. Corrige esto y vuelve a presionar el botón:" & vbCrLf & vbCrLf & errores, vbExclamation, TITULO
        Exit Sub
    End If

    ' ---------- 2. leer VENTAS (columnas B, D y K:M)
    nFilas = ULTIMA_FILA_VENTAS - 1
    vB = wsV.Range("B2:B" & ULTIMA_FILA_VENTAS).Value
    vD = wsV.Range("D2:D" & ULTIMA_FILA_VENTAS).Value
    vK = wsV.Range("K2:M" & ULTIMA_FILA_VENTAS).Value
    ReDim borrar(1 To nFilas)
    ReDim ocupada(1 To nFilas)
    For i = 1 To 10
        filaDest(i) = 0
    Next i

    ' ---------- 3. que filas de VENTAS ya pertenecen a esta factura
    For idx = 1 To nFilas
        mia = False
        If UCase(Trim(CStr(vK(idx, 1)))) = num Then mia = True
        If Trim(CStr(vK(idx, 3))) = ws.Name Then mia = True
        If mia Then
            llave = Trim(CStr(vK(idx, 2)))
            j = 0
            For i = 1 To 10
                If usa(i) Then
                    If llave = num & "|" & CStr(i) Then j = i
                End If
            Next i
            If j > 0 And filaDest(j) = 0 Then
                filaDest(j) = idx
                ocupada(idx) = True
            Else
                borrar(idx) = True
            End If
        End If
    Next idx

    ' ---------- 4. revisar existencias (descontando lo que esta factura ya habia registrado)
    For i = 1 To 10
        If usa(i) Then
            necesita = 0
            For j = 1 To 10
                If usa(j) Then
                    If UCase(cod(j)) = UCase(cod(i)) Then necesita = necesita + cant(j)
                End If
            Next j
            previa = 0
            For idx = 1 To nFilas
                If ocupada(idx) Or borrar(idx) Then
                    If UCase(Trim(CStr(vB(idx, 1)))) = UCase(cod(i)) Then
                        If IsNumeric(vD(idx, 1)) Then previa = previa + CDbl(vD(idx, 1))
                    End If
                End If
            Next idx
            hay = 0
            If IsNumeric(wsP.Cells(filaP(i), 9).Value) Then hay = CDbl(wsP.Cells(filaP(i), 9).Value)
            hay = hay + previa
            ' avisar una sola vez por codigo (en su primera aparicion)
            j = 0
            For k = 1 To i - 1
                If usa(k) Then
                    If UCase(cod(k)) = UCase(cod(i)) Then j = k
                End If
            Next k
            If j = 0 And necesita > hay Then
                avisos = avisos & "- " & cod(i) & " (" & TextoCelda(wsP.Cells(filaP(i), 2)) & "): hay " & CStr(hay) & " y la factura pide " & CStr(necesita) & "." & vbCrLf
            End If
        End If
    Next i
    If avisos <> "" Then
        If MsgBox("El inventario no alcanza para:" & vbCrLf & vbCrLf & avisos & vbCrLf & _
                  "¿Registrar la venta de todos modos (el stock quedará negativo)?", vbYesNo + vbExclamation + vbDefaultButton2, TITULO) <> vbYes Then
            Exit Sub
        End If
    End If

    ' ---------- 5. filas libres para las lineas nuevas
    idx = 1
    For i = 1 To 10
        If usa(i) And filaDest(i) = 0 Then
            Do While idx <= nFilas
                If Not ocupada(idx) Then
                    If borrar(idx) Then Exit Do
                    If Trim(CStr(vB(idx, 1))) = "" And Trim(CStr(vK(idx, 1))) = "" And Trim(CStr(vK(idx, 2))) = "" Then Exit Do
                End If
                idx = idx + 1
            Loop
            If idx > nFilas Then
                MsgBox "La hoja VENTAS está llena (llegó a la fila " & ULTIMA_FILA_VENTAS & "). No se registró nada.", vbExclamation, TITULO
                Exit Sub
            End If
            filaDest(i) = idx
            ocupada(idx) = True
            borrar(idx) = False
            nNuevas = nNuevas + 1
            idx = idx + 1
        End If
    Next i

    ' ---------- 6. medio de pago (ultimo abono con metodo; si no, POR CONFIRMAR)
    medio = "POR CONFIRMAR"
    For r = 28 To 24 Step -1
        If TextoCelda(ws.Cells(r, 4)) <> "" Then
            medio = UCase(TextoCelda(ws.Cells(r, 4)))
            Exit For
        End If
    Next r

    ' ---------- 7. escribir (sin recalcular a cada celda)
    Application.ScreenUpdating = False
    calc = Application.Calculation
    Application.Calculation = xlCalculationManual
    hechoCalc = True

    ' borrar movimientos viejos de esta factura que ya no corresponden
    For idx = 1 To nFilas
        If borrar(idx) Then
            r = idx + 1
            wsV.Cells(r, 1).ClearContents
            wsV.Cells(r, 2).ClearContents
            wsV.Cells(r, 4).ClearContents
            wsV.Cells(r, 6).ClearContents
            wsV.Cells(r, 10).ClearContents
            wsV.Cells(r, 11).ClearContents
            wsV.Cells(r, 12).ClearContents
            wsV.Cells(r, 13).ClearContents
            nBorr = nBorr + 1
        End If
    Next idx

    unidades = 0
    For i = 1 To 10
        If usa(i) Then
            r = filaDest(i) + 1
            If Trim(CStr(wsV.Cells(r, 11).Value)) <> "" Then nAct = nAct + 1
            wsV.Cells(r, 1).Value = fecha
            wsV.Cells(r, 1).NumberFormat = "mm-dd-yy"
            wsV.Cells(r, 2).Value = cod(i)
            wsV.Cells(r, 4).Value = cant(i)
            wsV.Cells(r, 6).Value = Round(cant(i) * (lista(i) - pre(i)), 2)
            wsV.Cells(r, 10).Value = medio
            wsV.Cells(r, 11).Value = num
            wsV.Cells(r, 12).Value = num & "|" & CStr(i)
            wsV.Cells(r, 13).Value = ws.Name
            unidades = unidades + cant(i)
        End If
    Next i

    AgregarFilaIndice ws

    Application.Calculation = calc
    hechoCalc = False
    Application.Calculate
    Application.ScreenUpdating = True

    If nLineas = 0 Then
        MsgBox "La factura " & num & " no tiene productos. Se quitaron sus movimientos de VENTAS (" & CStr(nBorr) & ").", vbInformation, TITULO
    ElseIf nAct = 0 And nBorr = 0 Then
        MsgBox "Venta " & num & " registrada." & vbCrLf & CStr(nLineas) & " producto(s), " & CStr(unidades) & " unidad(es) descontadas del inventario.", vbInformation, TITULO
    Else
        MsgBox "Venta " & num & " actualizada sin duplicar." & vbCrLf & _
               "Líneas nuevas: " & CStr(nNuevas) & "  |  Actualizadas: " & CStr(nAct) & "  |  Eliminadas: " & CStr(nBorr) & vbCrLf & _
               "Estado en ventas: " & TextoCelda(ws.Range("H10")), vbInformation, TITULO
    End If
    Exit Sub

Falla:
    Application.ScreenUpdating = True
    If hechoCalc Then
        On Error Resume Next
        Application.Calculation = calc
        On Error GoTo 0
    End If
    MsgBox "No se pudo registrar la venta: " & Err.Description & vbCrLf & _
           "Revisa la factura y vuelve a intentarlo. Si ya se escribió algo, el botón lo corrige sin duplicar.", vbCritical, TITULO
End Sub
