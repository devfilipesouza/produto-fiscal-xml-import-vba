Sub ImportarDadosNotaFiscal()
    ' === DECLARAÇÃO DE VARIÁVEIS ===
    Dim ws As Worksheet                       ' Planilha de destino ("CPV")
    Dim caminhoPasta As String                ' Caminho da pasta onde os XMLs estão armazenados
    Dim chaveAcesso As String                 ' Chave de acesso da nota fiscal (usada para localizar o arquivo)
    Dim caminhoArquivo As String              ' Caminho completo do arquivo XML
    Dim xmlDoc As Object                      ' Objeto DOM para manipular XML
    Dim xmlNode As Object                     ' Nó genérico do XML usado em loops
    Dim tempNode As Object                    ' Nó temporário usado para extrair valores
    Dim baseICMSST As String                  ' Base de cálculo do ICMS ST
    Dim valorICMSST As String                 ' Valor do ICMS ST
    Dim itens As Object                       ' Lista de nós <det> (itens da nota)
    Dim item As Object                        ' Nó de item individual
    Dim linha As Long                         ' Linha de escrita na planilha (inicia em 19)
    Dim totalQtd As Double                    ' Soma da quantidade total dos produtos
    Dim qComNode As Object                    ' Nó da quantidade comercial
    Dim noICMS As Object, noIPI As Object     ' Nós contendo informações de ICMS e IPI
    Dim aliqICMS As String, aliqIPI As String ' Alíquotas de ICMS e IPI
    Dim ultimaLinha As Long                   ' Última linha com dados na coluna F (para limpar)

    ' === INICIALIZA OBJETO E CAMINHO ===
    Set ws = ThisWorkbook.Sheets("CPV")
    caminhoPasta = "\\SRV-RELUZ\Users\ACESSO INTERNO\DOCUMENTOS FISCAIS\XML ENTRADA\"
    chaveAcesso = Trim(ws.Range("B12").MergeArea.Cells(1, 1).Value) ' Considera célula mesclada

    ' === VALIDA CHAVE DE ACESSO ===
    If chaveAcesso = "" Then
        MsgBox "Chave de acesso não informada. Por favor insira a chave de acesso.", vbExclamation
        Exit Sub
    End If

    ' === LIMPA CAMPOS DE DADOS ANTERIORES ===
    ws.Range("H12:Q12").ClearContents
    ws.Range("N16:P16").ClearContents
    ws.Range("B20:C20").ClearContents

    ' Limpa colunas F a K e N a O a partir da linha 19 (mantém L e M)
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

    ' === VERIFICA EXISTÊNCIA DO XML ===
    caminhoArquivo = caminhoPasta & chaveAcesso & ".xml"
    If Dir(caminhoArquivo) = "" Then
        MsgBox "Arquivo XML não encontrado para a chave informada.", vbCritical
        Exit Sub
    End If

    ' === CARREGA XML NA MEMÓRIA ===
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.Async = False
    xmlDoc.ValidateOnParse = False
    xmlDoc.Load (caminhoArquivo)

    ' Validação de erro ao carregar XML
    If xmlDoc.ParseError.ErrorCode <> 0 Then
        MsgBox "Erro ao carregar o XML: " & xmlDoc.ParseError.Reason, vbCritical
        Exit Sub
    End If

    ' Define namespace para XPath funcionar corretamente com NFe
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:nfe='http://www.portalfiscal.inf.br/nfe'"

    ' === CAMPOS PRINCIPAIS DA NOTA FISCAL ===
    ws.Range("H12").Value = xmlDoc.SelectSingleNode("//nfe:ide/nfe:nNF").Text           ' Nº da NF
    ws.Range("I12").Value = xmlDoc.SelectSingleNode("//nfe:ide/nfe:serie").Text         ' Série
    ws.Range("K12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vProd").Text ' Total dos produtos
    ws.Range("L12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vNF").Text   ' Valor total da nota

    ' === ICMS ST (Base e Valor) ===
    baseICMSST = ""
    valorICMSST = ""
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

    ' === OUTROS CAMPOS FISCAIS ===
    ws.Range("O12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vICMS").Text
    ws.Range("P12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vIPI").Text
    ws.Range("Q12").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vFrete").Text
    ws.Range("N16").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vOutro").Text
    ws.Range("O16").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vPIS").Text
    ws.Range("P16").Value = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vCOFINS").Text

    ' === LOOP PARA IMPORTAÇÃO DOS ITENS ===
    Set itens = xmlDoc.SelectNodes("//nfe:det")
    linha = 19
    totalQtd = 0

    For Each item In itens
        On Error Resume Next ' Em caso de campos ausentes

        ' Nome do produto
        ws.Range("F" & linha).Value = item.SelectSingleNode("nfe:prod/nfe:xProd").Text
        ' Valor unitário
        ws.Range("K" & linha).Value = item.SelectSingleNode("nfe:prod/nfe:vUnCom").Text

        ' Quantidade comercial (usada para somar total de itens)
        Set qComNode = item.SelectSingleNode("nfe:prod/nfe:qCom")
        If Not qComNode Is Nothing Then
            If IsNumeric(qComNode.Text) Then
                totalQtd = totalQtd + CDbl(qComNode.Text)
            End If
        End If

        ' Alíquota ICMS
        Set noICMS = item.SelectSingleNode("nfe:imposto/nfe:ICMS/*")
        If Not noICMS Is Nothing Then
            aliqICMS = noICMS.SelectSingleNode("nfe:pICMS").Text
            If IsNumeric(aliqICMS) Then
                ws.Range("N" & linha).Value = CDbl(aliqICMS) / 1000000 ' Convertendo p/ percentual
            End If
        End If

        ' Alíquota IPI
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

    ' === FINALIZAÇÃO ===
    MsgBox "Importação concluída com sucesso!", vbInformation

End Sub
