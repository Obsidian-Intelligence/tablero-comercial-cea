' ============================================================================
' Registro automático de eventos — Excel de Ventas CEA (versión VBA)
'
' Equivalente al Apps Script (apps_script/registro_eventos.gs) pero para un
' Excel de escritorio, sin depender de Google Sheets.
'
' INSTALACIÓN (una sola vez, sobre Excel_de_Ventas_CEA.xlsx):
'   1. Abrir el archivo en Excel de escritorio (Windows o Mac; no funciona
'      igual en Excel Online / la app web).
'   2. Guardar una copia como "Excel_de_Ventas_CEA.xlsm" (Libro de Excel
'      habilitado para macros). A partir de aquí se trabaja SOLO con el
'      .xlsm — el .xlsx normal no puede guardar macros.
'   3. Alt+F11 para abrir el editor de VBA.
'   4. Insertar → Módulo. Pegar ahí el bloque "MÓDULO ESTÁNDAR" de abajo.
'      Renombrar el módulo a "modEventos" (panel Propiedades, campo Name).
'   5. En el panel de proyecto (izquierda), doble clic en la hoja "Prospectos"
'      (bajo "Microsoft Excel Objetos"). Pegar ahí el bloque
'      "MÓDULO DE LA HOJA PROSPECTOS" de abajo — este NO va en un módulo
'      estándar, tiene que vivir en el objeto de esa hoja específica.
'   6. Doble clic en "ThisWorkbook" (mismo panel). Pegar el bloque
'      "MÓDULO ThisWorkbook" de abajo.
'   7. Guardar (Ctrl+S). Cerrar y volver a abrir el archivo: Excel va a pedir
'      "Habilitar contenido" / macros — hay que aceptarlo cada vez que se
'      abre, o marcar la carpeta como ubicación de confianza
'      (Archivo → Opciones → Centro de confianza → Ubicaciones de confianza).
'
' QUÉ HACE:
'   Cada vez que alguien cambia a mano la columna "etapa" en la hoja
'   Prospectos, se agrega automáticamente una fila en la hoja Eventos con:
'   id del prospecto, etapa anterior, etapa nueva, fecha/hora y el usuario
'   de Windows que hizo el cambio (Environ("Username") — no es un correo
'   como en la versión de Google, es el nombre de la cuenta de Windows con
'   la que se abrió Excel).
'
' LIMITACIONES vs. la versión de Google Sheets:
'   - Requiere que cada persona abra el archivo con macros habilitadas. Si
'     alguien las bloquea, sus cambios de etapa NO quedan registrados y no
'     hay aviso de que falló.
'   - Solo funciona en Excel de escritorio. Si el archivo se llega a abrir
'     y editar desde el navegador (Excel Online / OneDrive web), las macros
'     no corren.
'   - El identificador de quién editó es el usuario de Windows, no un
'     correo — pídele a cada persona que abra Excel con su propia sesión de
'     Windows/perfil, no una cuenta compartida.
'   - Si dos personas editan el archivo al mismo tiempo desde copias
'     distintas (por ejemplo, cada quien con su propia copia local
'     sincronizada por OneDrive/Drive), el caché de "etapa anterior" de cada
'     copia se puede desincronizar. Este método asume UN solo archivo que
'     todos abren y editan por turnos, no ediciones simultáneas.
' ============================================================================


' ---------------------------------------------------------------------------
' MÓDULO ESTÁNDAR (Insertar → Módulo, renombrar a "modEventos")
' ---------------------------------------------------------------------------

Public dictEtapas As Object

Sub CargarCacheEtapas()
    ' Lee toda la columna "etapa" de Prospectos a memoria, para poder
    ' comparar "antes" vs "después" cuando alguien edite una celda —
    ' Excel, a diferencia de Google Apps Script, no da el valor anterior
    ' de la celda de forma nativa.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Prospectos")

    Set dictEtapas = CreateObject("Scripting.Dictionary")

    Dim colId As Integer, colEtapa As Integer
    colId = ObtenerColumna(ws, "id")
    colEtapa = ObtenerColumna(ws, "etapa")
    If colId = 0 Or colEtapa = 0 Then Exit Sub

    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, colId).End(xlUp).Row

    Dim i As Long
    For i = 2 To ultimaFila
        Dim idProspecto As String
        idProspecto = CStr(ws.Cells(i, colId).Value)
        If idProspecto <> "" Then
            dictEtapas(idProspecto) = CStr(ws.Cells(i, colEtapa).Value)
        End If
    Next i
End Sub

Function ObtenerColumna(ws As Worksheet, nombreEncabezado As String) As Integer
    Dim c As Integer
    Dim ultimaCol As Integer
    ultimaCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To ultimaCol
        If ws.Cells(1, c).Value = nombreEncabezado Then
            ObtenerColumna = c
            Exit Function
        End If
    Next c
    ObtenerColumna = 0
End Function


' ---------------------------------------------------------------------------
' MÓDULO DE LA HOJA PROSPECTOS (doble clic en "Prospectos" en el panel de
' proyecto — este código vive EN ESA HOJA, no en un módulo estándar)
' ---------------------------------------------------------------------------

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim colEtapa As Integer, colId As Integer
    colEtapa = ObtenerColumna(Me, "etapa")
    colId = ObtenerColumna(Me, "id")
    If colEtapa = 0 Or colId = 0 Then Exit Sub

    If Target.Column <> colEtapa Then Exit Sub
    If Target.Row = 1 Then Exit Sub          ' encabezado
    If Target.Cells.Count > 1 Then Exit Sub  ' pegado masivo: no registrar

    If dictEtapas Is Nothing Then CargarCacheEtapas

    Dim idProspecto As String
    idProspecto = CStr(Me.Cells(Target.Row, colId).Value)
    If idProspecto = "" Then Exit Sub

    Dim etapaAnterior As String
    etapaAnterior = ""
    If dictEtapas.Exists(idProspecto) Then etapaAnterior = dictEtapas(idProspecto)

    Dim etapaNueva As String
    etapaNueva = CStr(Target.Value)

    Dim wsEventos As Worksheet
    Set wsEventos = ThisWorkbook.Sheets("Eventos")

    Dim filaNueva As Long
    filaNueva = wsEventos.Cells(wsEventos.Rows.Count, 1).End(xlUp).Row + 1

    wsEventos.Cells(filaNueva, 1).Value = idProspecto
    wsEventos.Cells(filaNueva, 2).Value = etapaAnterior
    wsEventos.Cells(filaNueva, 3).Value = etapaNueva
    wsEventos.Cells(filaNueva, 4).Value = Now
    wsEventos.Cells(filaNueva, 5).Value = Environ("Username")

    dictEtapas(idProspecto) = etapaNueva
End Sub


' ---------------------------------------------------------------------------
' MÓDULO ThisWorkbook (doble clic en "ThisWorkbook" en el panel de proyecto)
' ---------------------------------------------------------------------------

Private Sub Workbook_Open()
    CargarCacheEtapas
End Sub
