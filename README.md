<p align="center">
  <img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python Badge"/>
  <img src="https://img.shields.io/badge/SAP%20GUI%20Scripting-0FAAFF?style=for-the-badge&logo=sap&logoColor=white" alt="SAP Badge"/>
  <img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas Badge"/>
  <img src="https://img.shields.io/badge/openpyxl-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" alt="Excel Badge"/>
  <img src="https://img.shields.io/badge/Tkinter-FFD43B?style=for-the-badge&logo=python&logoColor=black" alt="Tkinter Badge"/>
</p>

<p align="center">
  Automação em <strong>Python</strong> que integra o <strong>SAP GUI Scripting</strong> a uma planilha
  <strong>Excel</strong> para identificar, entre as Requisições de Compra (PRs) do tipo <strong>J</strong>,
  quais estão vinculadas a um <strong>Elemento PEP</strong> (Investimento).
</p>

---

## Visão Geral

Este é um projeto de estágio que automatiza uma verificação manual repetitiva no **SAP**: para cada PR do
tipo J listada em uma planilha Excel, o script abre a transação **ME53N** (Visualizar Requisição de Compra),
navega até a Ordem de Compra vinculada e verifica se existe um **Elemento PEP** preenchido — indicando que a
requisição é um **Investimento**. O resultado é gravado automaticamente em uma nova coluna **"Comentários"**
no Excel de saída.

Com ele, é possível:

- Conectar-se a uma sessão já aberta do SAP GUI via **Scripting Engine**;
- Ler a planilha de entrada e filtrar apenas as PRs cujo `Tipo_IO` termina em **"J"**;
- Abrir cada PR na ME53N, acessar a Ordem de Compra e checar o campo `COAS-PSPEL` (Elemento PEP);
- Classificar cada PR como **"Investimento"**, **vazio** (despesa) ou **"Erro"**;
- Acompanhar o progresso em uma interface gráfica simples (**Tkinter**), com barra de progresso, tempo
  restante estimado e botão de cancelamento;
- Salvar o resultado periodicamente (auto-save), evitando perda de dados em execuções longas.

---

## Tecnologias Utilizadas

- **Linguagem:** Python
- **Integração SAP:** `pywin32` (`win32com.client`) via **SAP GUI Scripting Engine**
- **Manipulação de dados:** `pandas`
- **Leitura/escrita de Excel preservando formatação:** `openpyxl`
- **Interface gráfica:** `tkinter` / `ttk`
- **Concorrência:** `threading` (processo SAP roda em thread separada da UI)
- **COM/Threading Windows:** `pythoncom`

---

## Estrutura do Projeto

```texto
projeto-tipoJ/
├── teste_conexaoSAP.py             # Script mínimo: valida a conexão com o SAP GUI Scripting
├── tipoJ.py                        # Prova de conceito: abre uma PR fixa e verifica o Elemento PEP
└── teste_tipoj_corrigido_1.py      # Script principal: processa a planilha inteira, com
                                     # tratamento de erros, auto-save e barra de progresso
```

### Descrição dos Arquivos

| Arquivo | Papel no Projeto |
| :--- | :--- |
| **`teste_conexaoSAP.py`** | Script de sanidade. Apenas conecta ao SAP GUI já aberto e confirma o sucesso da conexão via `GetObject("SAPGUI")`. Usado como primeiro teste antes de qualquer automação. |
| **`tipoJ.py`** | Protótipo inicial. Abre a ME53N, digita manualmente o número de uma PR fixa, navega até a Ordem de Compra e tenta ler o Elemento PEP — base para a lógica que depois foi generalizada. |
| **`teste_tipoj_corrigido_1.py`** | Versão final e funcional. Lê a planilha de entrada inteira, itera por todas as PRs tipo J, executa a navegação completa no SAP para cada uma, grava o resultado no Excel e exibe uma interface gráfica de progresso. |

---

## Como o Script Principal Funciona

O `teste_tipoj_corrigido_1.py` é organizado em funções especializadas:

#### Conexão e Suporte
* **`get_sap_session()`**: obtém a sessão ativa do SAP GUI Scripting, com mensagem de erro explicando os pré-requisitos caso falhe (SAP fechado, scripting desabilitado, etc.).
* **`wait_for_element(session, element_id, timeout)`**: tenta localizar um elemento da tela SAP repetidamente até um timeout, evitando falhas por lentidão de carregamento.
* **`fechar_modais(session)`**: fecha popups residuais (`wnd[1]`, `wnd[2]`...) que possam ter ficado abertos de uma iteração anterior.

#### Navegação SAP
* **`voltar_para_inicio_me53n(session)`** / **`resetar_me53n(session)`**: garante que a sessão volta para a tela inicial da ME53N antes de processar a próxima PR, usando F3 repetidos e reforço via redigitação da transação.
* **`abrir_pr(session, pr)`**: abre o popup "Selecionar documento" (Shift+F5), digita o número da PR e confirma.
* **`verificar_investimento(session)`**: navega até a Ordem de Compra vinculada, acessa a aba "Atribuições" e lê o campo do Elemento PEP, retornando `"Investimento"` ou `""`.

#### Processamento e Interface
* **`ProgressApp`**: classe da interface Tkinter — barra de progresso, tempo restante estimado, botão de cancelamento e botão de conclusão.
* **`run_process(app)`**: laço principal que lê o Excel linha a linha, filtra as PRs tipo J, chama as funções de navegação SAP para cada uma, grava o resultado na coluna "Comentários" e realiza auto-save periódico.

---

## Regras de Negócio

| Regra | Descrição |
| :--- | :--- |
| **Identificação Tipo J** | Uma linha é considerada PR tipo J quando o campo `Tipo_IO` termina em "J" precedido de dígitos (ex.: `540J`, `0001J`). |
| **Investimento** | A PR é classificada como `"Investimento"` quando o campo Elemento PEP (`COAS-PSPEL`) da Ordem de Compra vinculada está preenchido. |
| **Despesa** | Quando o Elemento PEP não é encontrado ou está vazio, a coluna "Comentários" é deixada em branco. |
| **Erro** | Falhas de navegação durante o processamento de uma PR são registradas como `"Erro"`, e o script tenta se recuperar antes de seguir para a próxima linha. |

---

## Configuração da Planilha Excel

### Entrada esperada

A planilha de entrada deve conter, no mínimo, as colunas:

| Coluna | Descrição |
| :--- | :--- |
| **`Tipo_IO`** | Código do tipo de ordem interna; usado para filtrar as PRs tipo J. |
| **`Nº PR`** | Número da Requisição de Compra a ser aberta no SAP. |
| **`Centro de Custo`** | Referência usada para posicionar a nova coluna "Comentários" logo em seguida. |

### Saída gerada

O script adiciona (ou reutiliza, se já existir) a coluna **"Comentários"** imediatamente após "Centro de Custo",
preenchendo-a com `"Investimento"`, vazio ou `"Erro"` para cada PR processada.

### Caminhos dos arquivos

Os caminhos de entrada e saída estão definidos diretamente no script e devem ser ajustados conforme o ambiente local:

```python
arquivo_entrada = r"C:\Users\SEU_USUARIO\Downloads\arquivo_entrada.xlsx"
arquivo_saida   = r"C:\Users\SEU_USUARIO\Downloads\arquivo_saida.xlsx"
```

---

## Como Executar o Projeto

### Pré-requisitos

* **Windows** com **SAP GUI** instalado, aberto e **logado** na sessão desejada.
* **SAP GUI Scripting habilitado**: `Options > Accessibility & Scripting > Scripting` (também precisa estar habilitado no servidor SAP).
* **Python 3.x** instalado.
* Bibliotecas Python necessárias:

```bash
pip install pywin32 pandas openpyxl
```

> `tkinter`, `threading` e `pythoncom` já acompanham a instalação padrão do Python no Windows (pythoncom vem com pywin32).

### Passos

#### 1. Testar a conexão com o SAP
Antes de rodar o processo completo, valide se o SAP Scripting está acessível:

```bash
python teste_conexaoSAP.py
```
Se a mensagem `"Conectado com sucesso!"` aparecer, a conexão está funcionando.

#### 2. Ajustar os caminhos da planilha
Edite `teste_tipoj_corrigido_1.py` e atualize as variáveis `arquivo_entrada` e `arquivo_saida` com os caminhos reais da sua planilha.

#### 3. Executar o processamento completo

```bash
python teste_tipoj_corrigido_1.py
```

Uma janela com barra de progresso será exibida, mostrando a PR sendo processada e o tempo restante estimado.
É possível cancelar a execução a qualquer momento — os dados processados até então já estarão salvos no arquivo de saída pelo mecanismo de auto-save.

---

## Histórico de Correções

Este script passou por uma rodada de correções em relação à versão de protótipo (`tipoJ.py`):

* **Timeout de espera**: aumentado de 2s para 10s ao localizar elementos SAP, reduzindo falsos negativos em servidores mais lentos.
* **Retorno de tela (F3) faltante**: a função de verificação de investimento não retornava à tela inicial da ME53N após checar a Ordem de Compra, fazendo com que a PR seguinte fosse aberta na tela errada. Corrigido com `voltar_para_inicio_me53n()`.
* **Duplicação da coluna "Comentários"**: cada execução criava uma nova coluna, deslocando as demais. Corrigido com verificação de existência antes da inserção.
* **Filtro de Tipo_IO mais robusto**: substituído regex rígido por uma checagem que aceita qualquer código terminado em "J" precedido de dígitos.
* **Sincronização de linha do Excel**: `linha_excel` passou a ser derivada diretamente do índice do pandas, evitando dessincronização quando linhas eram puladas.

---

## Autor

Autor: Mariane S. Moura e João Mateus E. B. da Silva

Projeto desenvolvido e idealizado durante o estágio na Mercedes.
