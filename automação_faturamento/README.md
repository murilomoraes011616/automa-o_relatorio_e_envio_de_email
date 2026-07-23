# Automação de Relatório de Faturamento (SAP → Excel → PDF → Envio)

Automação em Python para o relatório semanal de faturamento/cobrança: puxa dados do SAP dentro do Excel, gera um PDF do relatório e envia por e-mail (Outlook) ou WhatsApp, eliminando o processo manual.

## Fluxo geral

```
Excel (.xlsm conectado ao SAP HANA)
  → Atualiza data de referência (célula X41, aba "tabelas-auxiliares")
  → RefreshAll() nas conexões SAP
  → Exporta aba "mapa_diario" como PDF (ExportAsFixedFormat)
  → Envia o PDF por:
      - Outlook (via COM/pywin32), ou
      - WhatsApp Web (via Playwright, sessão salva)
```

## Status atual

- ✅ **Geração do PDF**: primeira metade do fluxo (Excel → atualização SAP → exportação em PDF) está completa e funcionando.
- ✅ **Envio via WhatsApp**: versão funcional ponta a ponta usando Playwright — abre o `web.whatsapp.com` com sessão salva, localiza o contato pelo nome, anexa e envia o PDF.
- 🚧 **Envio via Outlook**: em construção — abrir o Outlook via COM/pywin32, montar e-mail (destinatários, cópia, assunto, corpo), anexar o PDF gerado e enviar.

## Decisões de projeto

- **Sem Selenium e sem automação de mouse/teclado** — a integração com Outlook usa exclusivamente APIs oficiais do Windows/Outlook (COM via `pywin32`).
- **Código modular e reutilizável**, pensado para servir de base para futuras automações, não só este relatório.

## Estrutura de arquivos planejada

```
main.py       # orquestra o fluxo (Excel → PDF → envio)
excel.py      # abertura do workbook, atualização SAP, exportação em PDF
pdf.py        # manipulação/validação do PDF gerado
outlook.py    # montagem e envio do e-mail via COM
config.py     # destinatários, caminhos de arquivo, assuntos — configurável sem mexer no código
logs.py       # logging da execução
utils.py      # funções auxiliares compartilhadas
```

## Boas práticas em andamento

- Espera real do término do `RefreshAll()` (em vez de `time.sleep()` fixo).
- Tratamento de exceções (`try/except`) em cada etapa crítica.
- Fechamento seguro do Excel mesmo em caso de erro (evitar processos "fantasmas" do Excel em segundo plano).
- Configuração externa (destinatários, caminhos, assuntos) separada do código, via `config.py`.

## Próximos passos

- Finalizar a integração com Outlook (montagem e envio do e-mail).
- Padronizar logging entre as versões Outlook e WhatsApp.
- Consolidar as duas versões de envio (Outlook / WhatsApp) sob a mesma orquestração em `main.py`.