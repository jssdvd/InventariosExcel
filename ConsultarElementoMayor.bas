Attribute VB_Name = "ConsultarElementoMayor"
Sub ConsultarElementoMayor()

    Dim numeroInventario As String
    Dim texto As Variant
    Dim info As Object
    
    texto = Application.InputBox("Introduzca el número de inventario", "Número de inventario")
    
    If texto = False Then
        Exit Sub
    End If
    
    numeroInventario = Digitos(CStr(texto))
    
    If numeroInventario = "" Then
        With ThisWorkbook.Worksheets("CONSULTAR ELEMENTO").Shapes("NumeroCodigoBarras")
            .TextFrame.Characters.Font.Size = 12
            .TextFrame.Characters.Text = ChrW(9654) & " Haz click aqui " & ChrW(9664)
        End With
        Exit Sub
    End If
    
    ThisWorkbook.Worksheets("CONSULTAR ELEMENTO").Shapes("NumeroCodigoBarras").TextFrame.Characters.Font.Size = 15
    ThisWorkbook.Worksheets("CONSULTAR ELEMENTO").Shapes("NumeroCodigoBarras").TextFrame.Characters.Text = numeroInventario
    
    Set info = DescargarInfo(numeroInventario)

    ThisWorkbook.Worksheets("CONSULTAR ELEMENTO").Cells(6, 4).Value = info("marcaElemento")
    ThisWorkbook.Worksheets("CONSULTAR ELEMENTO").Cells(8, 4).Value = info("numeroSerial")
    ThisWorkbook.Worksheets("CONSULTAR ELEMENTO").Cells(10, 4).Value = info("nombreElemento")
    ThisWorkbook.Worksheets("CONSULTAR ELEMENTO").Cells(13, 4).Value = info("codigoEdificio") & info("codigoAula")
    ThisWorkbook.Worksheets("CONSULTAR ELEMENTO").Cells(16, 4).Value = info("codigoUnidad")
    ThisWorkbook.Worksheets("CONSULTAR ELEMENTO").Cells(18, 4).Value = info("nombreResponsable")
    ThisWorkbook.Worksheets("CONSULTAR ELEMENTO").Cells(20, 4).Value = info("numeroDocumento")
    
End Sub

Function Digitos(texto As String) As String

    Dim textoTemporal As String
    
    textoTemporal = ""

    For i = 1 To Len(texto)
        If Mid(texto, i, 1) >= "0" And Mid(texto, i, 1) <= "9" Then
            textoTemporal = textoTemporal + Mid(texto, i, 1)
        End If
    Next

    Digitos = textoTemporal
    
End Function

Function DescargarInfo(numeroInventario As String) As Object

    Dim objetoPeticion As Object
    Dim url As String
    Dim respuesta As String

    Set objetoPeticion = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    url = ELEMENTOS_MAYORES_API & "num=" & numeroInventario & "&info&image"

    With objetoPeticion
        .Open "GET", url, False
        .setRequestHeader "Content-Type", "application/json"
        .send
        
        While objetoPeticion.readyState <> 4
            DoEvents
        Wend
        
        respuesta = .responseText

    End With

    On Error Resume Next
    Set DescargarInfo = JsonConverter.ParseJson(respuesta)
End Function
