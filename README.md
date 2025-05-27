# 📦 Importador de Produtos e valores Fiscais via XML (VBA)

Este projeto em VBA realiza a importação automatizada de produtos de Notas Fiscais Eletrônicas (NF-e) diretamente de arquivos XML para uma planilha do Excel. A automação tem como objetivo acelerar o processo de conferência, precificação e controle de estoque a partir de dados fiscais oficiais.

## ✅ Funcionalidades

- 📁 Leitura automática de arquivos XML de NF-e a partir de uma pasta de rede
- 🧾 Importação dos produtos listados nas notas fiscais
- 🔄 Mapeamento dos seguintes dados por produto:
  - Código do produto
  - Descrição
  - NCM
  - CFOP
  - Unidade
  - Quantidade
  - Valor unitário e total
- 📌 Relacionamento com a chave de acesso da NF-e
- 📊 Preenchimento automatizado em linhas organizadas por produto

## 📋 Estrutura Esperada da Planilha

A macro atua sobre a planilha **`PRNF`**, onde cada produto de uma NF-e ocupa uma linha individual. Os dados são organizados em colunas como:

| Coluna | Informação                      |
|--------|---------------------------------|
| J      | Chave de Acesso (NF-e)          |
| A~I    | Dados diversos do produto       |
| K      | Status de processamento         |

> ℹ️ As chaves de acesso devem estar preenchidas a partir da linha 13 na coluna J.

## 📁 Caminho dos Arquivos XML

A macro busca os arquivos XML da seguinte pasta de rede

