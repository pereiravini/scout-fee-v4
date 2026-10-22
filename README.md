# 📊 Scout Fee V4 — Automação e Conciliação Financeira de ERP em VBA

![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-Automation-blue?style=for-the-badge)
![ERP Benner](https://img.shields.io/badge/ERP-Benner-orange?style=for-the-badge)
![Process Automation](https://img.shields.io/badge/Focus-Process_Optimization-green?style=for-the-badge)

---

## 🔒 Aviso de Proteção de Dados & Privacidade (LGPD)

> **⚠️ Importante:** Todos os dados exibidos nas demonstrações, imagens, exemplos e na planilha-modelo de testes **são 100% fictícios**. Os nomes, valores, taxas, códigos R-LOCK e registros de fornecedores foram gerados manualmente exclusivamente para fins de portfólio e demonstração técnica. **Nenhuma informação confidencial, sensível ou real de clientes/terceiros é compartilhada neste repositório.**

---

## 🎯 Visão Geral

O **Scout Fee V4** é uma solução de automação e auditoria desenvolvida em VBA (Visual Basic for Applications) no Microsoft Excel para solucionar gargalos críticos na reconciliação e apuração de taxas de agência (*fee*) sobre vendas e eventos corporativos.

A ferramenta automatiza o cruzamento e a validação de dados entre o **Relatório de Produtividade** e o **Relatório de Registros Fee** exportados do ERP **Benner**, aplicando regras condicionais avançadas para identificar inconsistências instantaneamente.

> ⚡ **Impacto de Negócio:** Redução do tempo de auditoria diária de dezenas de minutos por projeto para **poucos segundos com um único clique**, eliminando a necessidade de conferência manual linha a linha em carteiras de eventos complexas.

---

## 🎬 Demonstração em Vídeo

> 📺 **[Clique aqui para assistir ao vídeo de demonstração da automação](./media/demo_Scout_Fee_2x.mp4)**
> 
> ⏩ *Demonstração acelerada em 2x cobrindo a carga de dados brutos do ERP Benner, higienização automática, aplicação do motor de regras lógicas e diagnóstico de divergências.*

---

## 🛠️ O Problema vs. A Solução

### ❌ O Cenário Anterior (Manual)
* **Volume & Complexidade:** Projetos com até 300+ lançamentos de vendas (RLOCs) realizados por múltiplos operadores, sujeitos a inconsistências cadastrais.
* **Pontos de Falha Humana:** Lançamentos com *fee* zerado, aplicação indevida de taxas em custos internos/aéreos, duplicidades ou faturamento contra cliente incorreto.
* **Gargalo Operacional:** Na ocorrência de divergência no valor final do *fee*, o analista precisava auditar manualmente todos os registros no ERP sem uma regra clara de diagnóstico, gerando atrito e atraso no fechamento.

### ✅ A Solução com o Scout Fee V4
* **Cruzamento Automatizado:** A planilha consome as extrações brutas do ERP Benner e executa um motor de validação lógica em VBA.
* **Rastreabilidade total:** Atua como um "diário de bordo" do fechamento, permitindo consultar e auditar o histórico de status e divergências de meses anteriores.
* **Diagnóstico de Causa-Raiz:** O sistema aponta exatamente em qual RLOC ocorreu o problema e qual a causa (ex: taxa em custo interno, duplicidade), direcionando a correção no ERP Benner de forma cirúrgica.

---

## 🚀 Principais Funcionalidades

### 1. 🧹 Higienização e Formatação Dinâmica
* Padronização estrutural e visual dos relatórios brutos do ERP Benner com apenas um clique.
* Remoção de colunas desnecessárias e ajuste automático de layout.

### 2. 🧠 Motor de Análise PP (Conciliação & Validações)
* **Caixa Interativa:** Pergunta dinâmica ao usuário para aplicar/isentar regras de *Fee Aéreo* apurado por fora.
* **Fórmulas e Injeção de Regras:** Inserção automática de colunas calculadas para interceptar divergências de valores, duplicidades ou tributações indevidas.
* **Painel Dinâmico de Fee:** Associação automática de fornecedores e cálculo de saldo divergente (*Fee Ideal vs. Fee Apurado*), atualizando dinamicamente conforme os lançamentos são corrigidos.

### 3. ➕ Carga Incremental (*Novos Registros*)
* Permite importar novos relatórios de produtividade sem perder as anotações e status já preenchidos.
* Executa *deduplicação* por chave única (RLOC) e insere apenas os novos lançamentos, reportando a quantidade exata de linhas adicionadas.

### 4. 📤 Preparação para Carga no ERP (*Fee por Arquivo*)
* Processamento e limpeza de dados para geração de arquivos no formato exigido para importação no ERP Benner.
* Elimina duplicatas e valida registros garantindo a integridade dos dados na reimportação.

---

## 💡 Competências Transferíveis & Visão de ERP

Embora esta automação tenha sido customizada para otimizar os fluxos do **ERP Benner**, a arquitetura e a lógica aplicadas são **100% transferíveis para qualquer ERP corporativo** (como SAP, TOTVS Protheus, Oracle, Senior ou Sankhya) ou bancos de dados relacionais:

* **Mapeamento de Regras de Negócio:** Tradução de regras financeiras/comerciais complexas em algoritmos de validação automática.
* **Tratamento Agnóstico de Dados:** Estrutura preparada para consumir, higienizar e cruzar extratos/relatórios brutos vindos de qualquer sistema de gestão.
* **Cultura Problem-Solver (Low-Resource / High-Impact):** Capacidade de gerar alto ganho de produtividade e redução de tempo operacional utilizando recursos já disponíveis (Excel/VBA).

---

## 📋 Pré-requisitos e Como Utilizar

1. **Requisitos:** Microsoft Excel 2016 ou superior com suporte a Macros habilitado (`.xlsm`).
2. **Passo a Passo:**
   1. Abra a planilha `Scout_Fee_V4_Template.xlsm`.
   2. Cole o relatório de produtividade do ERP Benner na aba `Produtividade`.
   3. Clique no botão **"Formatar"**.
   4. Cole o relatório de fee na aba `Registros Fee`.
   5. Retorne à aba de produtividade e clique em **"Análise PP"**.
   6. Responda à verificação de *Fee Aéreo* na tela e acompanhe a sinalização automática das divergências.

💡 Arquivos de Teste: Dentro da pasta data/, você encontra exemplos de relatórios brutos do ERP Benner com dados fictícios para testar as macros na prática.

---

## ✒️ Autor

Desenvolvido por **Vinicius Pereira**  
*Analista de Dados & Automação de Processos*  
[LinkedIn](https://www.linkedin.com/in/vinicius-pereira-a92689172) • [GitHub](https://github.com/pereiravini)