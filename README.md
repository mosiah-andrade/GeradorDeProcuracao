
# ☀️ Gerador de Procurações - Homolog Solar

Este projeto é uma ferramenta de automação desenvolvida em Python para gerar procurações personalizadas (em **DOCX** e **PDF**) a partir de dados inseridos em uma planilha Excel.

O sistema identifica a concessionária e o tipo de representante (CPF ou CNPJ), seleciona o modelo de contrato correto, preenche as variáveis e converte o resultado final para PDF utilizando o Microsoft Word.

---

## 🚀 Funcionalidades

- **Leitura de Dados:** Extrai informações automaticamente de uma planilha Excel (`.xlsx` ou `.xlsm`).
- **Seleção Inteligente:** Escolhe o modelo de procuração correto com base na Concessionária (Celpe, Coelba, Cosern, Equatorial) e no tipo de cliente (Pessoa Física ou Jurídica).
- **Preenchimento Automático:** Substitui marcadores (ex: `{{NOME}}`, `{{CPF}}`) pelos dados reais.
- **Geração de PDF:** Converte o documento final para PDF automaticamente.
- **Independência de Pastas:** O Excel pode estar em qualquer lugar; o sistema salva os contratos gerados na mesma pasta da planilha.

---

## 🛠️ Pré-requisitos

Para executar o código fonte ou garantir que o executável funcione corretamente, é necessário:

1.  **Microsoft Word Instalado:** O script utiliza o Word instalado na máquina para garantir uma conversão perfeita para PDF.
2.  **Sistema Operacional Windows:** Devido à dependência do Microsoft Word (COM Interface).
3.  **Python 3.10+** (Apenas se for rodar o script diretamente).

### 📚 Bibliotecas Python Necessárias

Se for rodar pelo código fonte, instale as dependências:

```bash
pip install openpyxl python-docx docx2pdf

```

## 📂 Estrutura de Pastas Obrigatória
Para que o sistema (seja o script `.py` ou o `.exe`) encontre os modelos, a estrutura de pastas deve ser mantida **exatamente** como abaixo:

```
📁 Pasta do Sistema (C:\SistemaHomolog\ ou similar)
│
├── 📜 GeradorProcuracao.exe      (O Executável)
│
├── 📂 Procuração-celpe
│   ├── MODELO-PROCURAÇÃO-Celpe-CPF.docx
│   └── MODELO-PROCURAÇÃO-Celpe-CNPJ.docx
│
├── 📂 Procuração-Coelba
│   ├── MODELO-PROCURAÇÃO-Coelba-CPF.docx
│   └── MODELO-PROCURAÇÃO-Coelba-CNPJ.docx
│
├── 📂 Procuração-Cosern
│   ├── MODELO-PROCURAÇÃO-Cosern-CPF.docx
│   └── MODELO-PROCURAÇÃO-Cosern-CNPJ.docx
│
└── 📂 Procuração-Equatorial
    ├── MODELO-PROCURAÇÃO-Equatorial-CPF.docx
    └── MODELO-PROCURAÇÃO-Equatorial-CNPJ.docx

```

**Nota:** A planilha `DADOS_DO_CLIENTE.xlsx` pode ficar em qualquer outra pasta (ex: Área de Trabalho, Documentos). O executável deve ficar fixo junto com as pastas dos modelos.

## 📦 Como Gerar o Executável (.exe)

Para transformar o script Python em um programa executável que funciona em outros computadores (desde que tenham o Word instalado), use o **PyInstaller**.

1. Abra o terminal na pasta do script.

2. Execute o comando:

```bash
python -m PyInstaller --onefile --name "GeradorProcuracao" gerar_contrato.py
```
3. O arquivo GeradorProcuracao.exe será criado na pasta dist. Mova-o para a "Pasta do Sistema" junto com as pastas dos modelos.

## 🖥️ Integração com Excel (VBA)

Para chamar este gerador através de um botão no Excel, utilize o seguinte código VBA no seu módulo:

```VBA
Sub ExecutarPython()
    Dim CaminhoExe As String
    Dim PlanilhaAtual As String
    Dim Comando As String
    
    ' 1. Salvar Planilha
    ThisWorkbook.Save
    
    ' 2. Caminho Fixo do Sistema (Onde você guardou o .exe e os modelos)
    CaminhoExe = "C:\SistemaHomolog\GeradorProcuracao.exe"
    
    ' 3. Caminho da Planilha (Enviado para o Python saber onde salvar)
    PlanilhaAtual = ThisWorkbook.FullName
    
    ' 4. Executa
    Comando = """" & CaminhoExe & """ """ & PlanilhaAtual & """"
    Call Shell(Comando, vbNormalFocus)
End Sub

```

## ⚠️ Solução de Problemas Comuns
1. Erro de Permissão (PermissionError):

    - Verifique se não há uma versão antiga do .exe rodando no Gerenciador de Tarefas.

    - Se o erro for ao ler o Excel, verifique se a planilha não está travada por outro usuário na rede.

2. Erro ao gerar PDF:

    - Certifique-se de que não há janelas de diálogo do Word abertas (como "Salvar como" ou ativação).

    - O Microsoft Word deve estar instalado e ativado na máquina.

3. Modelos não encontrados:

    - Confira se os nomes das pastas (Procuração-celpe, etc.) e os nomes dos arquivos .docx estão exatamente iguais aos descritos na seção "Estrutura de Pastas".

Desenvolvido para **Homolog Solar**.