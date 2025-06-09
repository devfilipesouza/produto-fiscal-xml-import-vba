Sub IMPORTAR_NOTA_FISCAL_CPV()

    ' Declaração das variáveis
    Dim ws As Worksheet ' Planilha de trabalho
    Dim caminhoPasta As String ' Caminho onde estão os arquivos XML
    Dim chaveAcesso As String ' Chave de acesso da nota fiscal
    Dim caminhoArquivo As String ' Caminho completo do arquivo XML
    Dim xmlDoc As Object ' Objeto XML para carregar o arquivo
    Dim xmlNode As Object, tempNode As Object ' Nós temporários de leitura do XML
    Dim baseICMSST As String, valorICMSST As String ' Valores fiscais ICMS ST
    Dim itens As Object, item As Object ' Lista de itens da nota
    Dim linha As Long ' Linha da planilha para escrever os itens
    Dim totalQtd As Double ' Total de quantidade dos produtos
    Dim qComNode As Object ' Nó da quantidade comercial
    Dim noICMS As Object, noIPI As Object ' Nós ICMS/IPI de cada item
    Dim aliqICMS As String, aliqIPI As String ' Alíquotas de ICMS/IPI
    Dim ultimaLinha As Long ' Última linha preenchida da tabela
    Dim encontrouICMS As Boolean ' Flag se encontrou ICMS
    Dim resposta As VbMsgBoxResult ' Resposta do usuário em uma MsgBox
    Dim xmlNodes As Object ' Lista de nós de ICMS para a nota

    ' Define a planilha "CPV"
    Set ws = ThisWorkbook.Sheets("CPV")
    
    ' Caminho fixo onde os arquivos XML estão localizados
    caminhoPasta = "\\SRV-RELUZ\Users\ACESSO INTERNO\DOCUMENTOS FISCAIS\XML ENTRADA\"

    ' Lê a chave de acesso da célula B12
    chaveAcesso = Trim(ws.Range("B12").MergeArea.Cells(1, 1).Value)

    ' Valida se a chave foi informada
    If chaveAcesso = "" Then
        MsgBox "Chave de acesso não informada. Por favor insira a chave de acesso.", vbExclamation
        Exit Sub
    End If

    ' Limpa os dados antigos da planilha
    ws.Range("H12:Q12").ClearContents
    ws.Range("N16:P16").ClearContents
    ws.Range("B20:C20").ClearContents
    ws.Range("C27").Value = 0

    ' Limpa a área da tabela de itens, se houver registros
    ultimaLinha = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    If ultimaLinha >= 19 Then
        ws.Range("F19:K" & ultimaLinha).ClearContents
        ws.Range("N19:O" & ultimaLinha).ClearContents
    End If

    ' Monta o caminho completo do XML com base na chave
    caminhoArquivo = caminhoPasta & chaveAcesso & ".xml"

    ' Verifica se o arquivo existe
    If Dir(caminhoArquivo) = "" Then
        MsgBox "Arquivo XML não encontrado para a chave informada.", vbCritical
        Exit Sub
    End If

    ' Cria e carrega o documento XML
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.Async = False
    xmlDoc.ValidateOnParse = False
    xmlDoc.Load (caminhoArquivo)

    ' Verifica se houve erro ao carregar
    If xmlDoc.ParseError.ErrorCode <> 0 Then
        MsgBox "Erro ao carregar o XML: " & xmlDoc.ParseError.Reason, vbCritical
        Exit Sub
    End If

    ' Define o namespace usado nas notas fiscais eletrônicas
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:nfe='http://www.portalfiscal.inf.br/nfe'"

    ' Captura dados principais da nota (número, série, valor dos produtos e valor total)
    On Error Resume Next ' Impede interrupção por erro
    ws.Range("H12").Value = xmlDoc.SelectSingleNode("//nfe:ide/nfe:nNF").Text
    ws.Range("I12").Value = xmlDoc.SelectSingleNode("//nfe:ide/nfe:serie").Text
    ws.Range("K12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vProd").Text
    ws.Range("L12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vNF").Text
    On Error GoTo 0 ' Retoma tratamento normal de erro

    ' ICMS ST – busca os campos vBCST (base) e vICMSST (valor)
    Set xmlNodes = xmlDoc.SelectNodes("//nfe:imposto/nfe:ICMS/*")
    baseICMSST = ""
    valorICMSST = ""

    If xmlNodes.Length = 0 Then
        ' Caso não tenha ICMS ST
        ws.Range("B20").Value = 1
        ws.Range("C20").Value = 0
        ws.Range("N12").Value = 0
        MsgBox "Nota fiscal não possui ICMS ST.", vbExclamation
    Else
        ' Percorre os nós de ICMS buscando as informações
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

        ' Preenche os campos na planilha com os valores encontrados
        ws.Range("B20").Value = IIf(baseICMSST = "", 1, baseICMSST)
        ws.Range("C20").Value = IIf(valorICMSST = "", 0, valorICMSST)
        ws.Range("N12").Value = IIf(valorICMSST = "", 0, valorICMSST)
    End If

    ' Outros campos fiscais (ICMS, IPI, Frete, PIS, COFINS)
    On Error Resume Next
    ws.Range("O12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vICMS").Text
    ws.Range("P12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vIPI").Text
    ws.Range("Q12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vFrete").Text
    ws.Range("N16").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vOutro").Text
    ws.Range("O16").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vPIS").Text
    ws.Range("P16").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vCOFINS").Text
    On Error GoTo 0

    ' Percorre os itens da nota fiscal
    Set itens = xmlDoc.SelectNodes("//nfe:det")
    linha = 19
    totalQtd = 0
    encontrouICMS = False

    For Each item In itens
        On Error Resume Next

        ' Nome e valor unitário do produto
        ws.Range("F" & linha).Value = item.SelectSingleNode("nfe:prod/nfe:xProd").Text
        ws.Range("K" & linha).Value = item.SelectSingleNode("nfe:prod/nfe:vUnCom").Text

        ' Soma a quantidade total dos itens
        Set qComNode = item.SelectSingleNode("nfe:prod/nfe:qCom")
        If Not qComNode Is Nothing Then
            If IsNumeric(qComNode.Text) Then totalQtd = totalQtd + CDbl(qComNode.Text)
        End If

        ' Busca alíquota do ICMS
        Set noICMS = item.SelectSingleNode("nfe:imposto/nfe:ICMS/*")
        If Not noICMS Is Nothing Then
            aliqICMS = noICMS.SelectSingleNode("nfe:pICMS").Text
            If IsNumeric(aliqICMS) Then
                ws.Range("N" & linha).Value = CDbl(aliqICMS) / 1000000
                encontrouICMS = True
            End If
        End If

        ' Busca alíquota do IPI
        Set noIPI = item.SelectSingleNode("nfe:imposto/nfe:IPI/nfe:IPITrib")
        If Not noIPI Is Nothing Then
            aliqIPI = noIPI.SelectSingleNode("nfe:pIPI").Text
            If IsNumeric(aliqIPI) Then ws.Range("O" & linha).Value = CDbl(aliqIPI) / 1000000
        End If

        On Error GoTo 0
        linha = linha + 1
    Next item

    ' Caso nenhum item tenha ICMS, pode ser uma nota de PE (regra interna)
    If Not encontrouICMS Then
        MsgBox "Nota fiscal não possui ICMS nos produtos.", vbExclamation
        resposta = MsgBox("A nota fiscal pertence ao estado de Pernambuco (PE)?", vbYesNo + vbQuestion, "Confirmação")
        If resposta = vbYes Then
            ws.Range("C27").Value = 0.205
        Else
            ws.Range("C27").Value = 0.04
        End If
    End If

    ' Finalização
    MsgBox "Importação concluída com sucesso!", vbInformation

End Sub
