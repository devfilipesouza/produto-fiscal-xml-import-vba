Sub ImportarICMSST()
    Dim ws As Worksheet
    Dim caminhoPasta As String
    Dim chaveAcesso As String
    Dim caminhoArquivo As String
    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim tempNode As Object
    Dim baseICMSST As String
    Dim valorICMSST As String
    Dim itens As Object
    Dim item As Object
    Dim linha As Long
    Dim totalQtd As Double
    Dim qComNode As Object
    Dim noICMS As Object, noIPI As Object
    Dim aliqICMS As String, aliqIPI As String
    Dim ultimaLinha As Long
    Dim encontrouICMS As Boolean
    Dim resposta As VbMsgBoxResult

    Set ws = ThisWorkbook.Sheets("CPV")
    caminhoPasta = "\\SRV-RELUZ\Users\ACESSO INTERNO\DOCUMENTOS FISCAIS\XML ENTRADA\"
    chaveAcesso = Trim(ws.Range("B12").MergeArea.Cells(1, 1).Value)

    If chaveAcesso = "" Then
        MsgBox "Chave de acesso não informada. Por favor insira a chave de acesso.", vbExclamation
        Exit Sub
    End If

    ' --- LIMPA CAMPOS ANTES DE IMPORTAR NOVOS DADOS ---
    ws.Range("H12:Q12").ClearContents
    ws.Range("N16:P16").ClearContents
    ws.Range("B20:C20").ClearContents
    ws.Range("C27").Value = 0 ' Zera C27 no início

    ultimaLinha = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    If ultimaLinha >= 19 Then
        ws.Range("F19:F" & ultimaLinha).ClearContents
        ws.Range("G19:G" & ultimaLinha).ClearContents
        ws.Range("H19:H" & ultimaLinha).ClearContents
        ws.Range("I19:I" & ultimaLinha).ClearContents
        ws.Range("J19:J" & ultimaLinha).ClearContents
        ws.Range("K19:K" & ultimaLinha).ClearContents
        ws.Range("N19:N" & ultimaLinha).ClearContents
        ws.Range("O19:O" & ultimaLinha).ClearContents
    End If

    caminhoArquivo = caminhoPasta & chaveAcesso & ".xml"

    If Dir(caminhoArquivo) = "" Then
        MsgBox "Arquivo XML não encontrado para a chave informada.", vbCritical
        Exit Sub
    End If

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.Async = False
    xmlDoc.ValidateOnParse = False
    xmlDoc.Load (caminhoArquivo)

    If xmlDoc.ParseError.ErrorCode <> 0 Then
        MsgBox "Erro ao carregar o XML: " & xmlDoc.ParseError.Reason, vbCritical
        Exit Sub
    End If

    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:nfe='http://www.portalfiscal.inf.br/nfe'"

    ' --- Informações principais da nota ---
    ws.Range("H12").Value = xmlDoc.SelectSingleNode("//nfe:ide/nfe:nNF").Text
    ws.Range("I12").Value = xmlDoc.SelectSingleNode("//nfe:ide/nfe:serie").Text
    ws.Range("K12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vProd").Text
    ws.Range("L12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vNF").Text

    ' --- ICMS ST ---
    baseICMSST = ""
    valorICMSST = ""

    Dim xmlNodes As Object
    Set xmlNodes = xmlDoc.SelectNodes("//nfe:ICMS/*[nfe:vBCST or nfe:vICMSST]")

    If xmlNodes.Length = 0 Then
        ws.Range("B20").Value = 1
        ws.Range("C20").Value = 0
        MsgBox "Nota fiscal não possui ICMS ST.", vbExclamation
    Else
        For Each xmlNode In xmlNodes
            If baseICMSST = "" Then
                Set tempNode = xmlNode.SelectSingleNode("nfe:vBCST")
                If Not tempNode Is Nothing Then baseICMSST = tempNode.Text
            End If
            If valorICMSST = "" Then
                Set tempNode = xmlNode.SelectSingleNode("nfe:vICMSST")
                If Not tempNode Is Nothing Then valorICMSST = tempNode.Text
            End If
        Next xmlNode

        ws.Range("B20").Value = IIf(baseICMSST = "", 1, baseICMSST)
        ws.Range("C20").Value = IIf(valorICMSST = "", 0, valorICMSST)
    End If

    ' --- Outros campos fiscais ---
    ws.Range("O12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vICMS").Text
    ws.Range("P12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vIPI").Text
    ws.Range("Q12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vFrete").Text
    ws.Range("N16").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vOutro").Text
    ws.Range("O16").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vPIS").Text
    ws.Range("P16").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vCOFINS").Text

    ' --- Itens da nota ---
    Set itens = xmlDoc.SelectNodes("//nfe:det")
    linha = 19
    totalQtd = 0
    encontrouICMS = False

    For Each item In itens
        On Error Resume Next

        ws.Range("F" & linha).Value = item.SelectSingleNode("nfe:prod/nfe:xProd").Text
        ws.Range("K" & linha).Value = item.SelectSingleNode("nfe:prod/nfe:vUnCom").Text

        Set qComNode = item.SelectSingleNode("nfe:prod/nfe:qCom")
        If Not qComNode Is Nothing Then
            If IsNumeric(qComNode.Text) Then
                totalQtd = totalQtd + CDbl(qComNode.Text)
            End If
        End If

        Set noICMS = item.SelectSingleNode("nfe:imposto/nfe:ICMS/*")
        If Not noICMS Is Nothing Then
            aliqICMS = noICMS.SelectSingleNode("nfe:pICMS").Text
            If IsNumeric(aliqICMS) Then
                ws.Range("N" & linha).Value = CDbl(aliqICMS) / 1000000
                encontrouICMS = True
            End If
        End If

        Set noIPI = item.SelectSingleNode("nfe:imposto/nfe:IPI/nfe:IPITrib")
        If Not noIPI Is Nothing Then
            aliqIPI = noIPI.SelectSingleNode("nfe:pIPI").Text
            If IsNumeric(aliqIPI) Then
                ws.Range("O" & linha).Value = CDbl(aliqIPI) / 1000000
            End If
        End If

        On Error GoTo 0
        linha = linha + 1
    Next item

    ' --- Regra especial para PE quando não há ICMS ---
    If Not encontrouICMS Then
    MsgBox "Nota fiscal não possui ICMS nos produtos.", vbExclamation
    resposta = MsgBox("A nota fiscal pertence ao estado de Pernambuco (PE)?", vbYesNo + vbQuestion, "Confirmação")
    If resposta = vbYes Then
        ws.Range("C27").Value = 0.205
    Else
        ws.Range("C27").Value = 0.04 ' valor alterado de 0 para 0.04
        End If
    End If


    MsgBox "Importação concluída com sucesso!", vbInformation
End Sub
