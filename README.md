# Automações em Python (xlwings + integrações)

Repositório de scripts de automação para eliminar tarefas manuais e repetitivas em Excel/relatórios corporativos, usando Python (principalmente `xlwings` para controlar o Excel de verdade, via COM).

## Projetos

### 1. [Automação de Relatório de Faturamento](./README-faturamento.md)
Puxa dados do SAP dentro de um Excel conectado (SAP HANA), gera um PDF do relatório semanal de faturamento/cobrança e envia automaticamente — por e-mail (Outlook) ou WhatsApp.

- Geração do PDF: **funcionando**
- Envio via WhatsApp (Playwright): **funcionando**
- Envio via Outlook (COM/pywin32): **em construção**

### 2. [KPI - Mapa de Vendas](./README-mapa-de-vendas.md)
Automatiza o ajuste do relatório "Mapa de Vendas": filtra pedidos com valores inconsistentes (`#N/D`, `0`, vazio, `#VALOR!`) e organiza esses casos numa aba separada para tratamento manual, sem precisar filtrar e copiar linha por linha na mão.

- Status: **em desenvolvimento**, usado também como projeto de estudo prático de xlwings.

## Stack comum

- **Python** + **xlwings** — controla uma instância real do Excel (não lê/escreve o arquivo "por fora", manipula a aplicação como se fosse um usuário).
- **pywin32 (COM)** — integração com Outlook, sem depender de automação de mouse/teclado ou Selenium.
- **Playwright** — automação do WhatsApp Web para envio de relatórios.

## Filosofia dos projetos

- Preferência por **APIs oficiais** (COM do Windows/Excel/Outlook) em vez de automação de interface (cliques simulados, Selenium).
- Código pensado para ser **modular e reutilizável** entre automações futuras, não só para resolver o caso específico de cada relatório.
- Tratamento de erro e logging como parte do padrão, não como extra.

## Como navegar este repositório

Cada projeto tem seu próprio README com detalhes de fluxo, estrutura de arquivos e status atual — os links acima levam direto pra cada um.