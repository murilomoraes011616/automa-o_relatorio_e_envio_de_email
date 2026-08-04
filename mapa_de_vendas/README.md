# 📊 Automação do Mapa de Vendas

Automação desenvolvida em **Python** para eliminar atividades manuais do processo diário de geração do **Mapa de Vendas**.

O projeto realiza desde a atualização dos dados no Excel até o envio automático do relatório em PDF para todos os destinatários da empresa utilizando o Microsoft Outlook.

---

# 🚀 Funcionalidades

O sistema executa automaticamente todo o fluxo abaixo:

* Atualiza as conexões do Excel e do Power Query.
* Filtra registros com erros.
* Trata e padroniza os dados.
* Classifica automaticamente contratos de Locação.
* Identifica pedidos que devem ser excluídos.
* Remove pedidos duplicados.
* Alimenta automaticamente a tabela utilizada pelo Power Query.
* Atualiza novamente todas as consultas.
* Configura a impressão do relatório.
* Exporta o Mapa de Vendas para PDF.
* Cria automaticamente um e-mail no Outlook.
* Insere assunto e corpo do e-mail.
* Carrega automaticamente a assinatura do Outlook.
* Adiciona destinatários e cópias.
* Anexa o PDF gerado.
* Abre o e-mail para conferência (ou pode ser configurado para envio automático).

---

# 🛠 Tecnologias utilizadas

* Python 3
* xlwings
* pywin32
* Microsoft Excel
* Microsoft Outlook
* Power Query

---

# 📂 Estrutura do projeto

```text
📁 automacao_mapa_vendas
│
├── main.py                 # Automação do Excel
├── envio_email.py          # Automação do Outlook
├── README.md
│
└── MAPA DE VENDAS.pdf      # Relatório gerado automaticamente
```

---

# 🔄 Fluxo completo da automação

```text
Início
   │
   ▼
Abre o Excel
   │
   ▼
Atualiza Power Query
   │
   ▼
Filtra registros
   │
   ▼
Corrige informações
   │
   ▼
Classifica contratos
   │
   ▼
Gera lista de exclusões
   │
   ▼
Atualiza novamente as consultas
   │
   ▼
Exporta o relatório em PDF
   │
   ▼
Abre o Outlook
   │
   ▼
Cria um novo e-mail
   │
   ▼
Preenche assunto
   │
   ▼
Preenche corpo da mensagem
   │
   ▼
Carrega assinatura do Outlook
   │
   ▼
Adiciona destinatários
   │
   ▼
Adiciona cópias
   │
   ▼
Anexa o PDF
   │
   ▼
Abre o e-mail para conferência
```

---

# 📈 Parte 1 — Tratamento do Excel

A primeira etapa da automação é responsável por todo o processamento dos dados.

## O que é realizado

* Atualização automática das conexões.
* Aplicação de filtros.
* Correção de classificações.
* Identificação de contratos.
* Geração da lista de PVS excluídos.
* Atualização do Power Query.
* Exportação automática do relatório em PDF.

Essa etapa elimina praticamente todo o trabalho manual realizado dentro do Excel.

---

# 📧 Parte 2 — Envio automático do relatório

Após a geração do PDF, uma segunda automação inicia utilizando a biblioteca **pywin32**, permitindo controlar o Microsoft Outlook através da interface COM do Windows.

O script executa automaticamente as seguintes etapas:

* conecta ao Outlook;
* cria um novo e-mail;
* obtém automaticamente a assinatura padrão do usuário;
* monta o corpo da mensagem em HTML;
* insere a data atual no texto do e-mail;
* define o assunto;
* adiciona todos os destinatários;
* adiciona os destinatários em cópia (CC);
* anexa o PDF recém-gerado;
* abre o e-mail para conferência antes do envio.

Caso desejado, o método `Display()` pode ser substituído por `Send()`, permitindo envio totalmente automático.

---

# 📬 Corpo do e-mail

O sistema gera automaticamente uma mensagem semelhante a:

> Bom dia,
>
> Seguem anexos o mapa de venda, atualizado até **DD/MM/AAAA**.
>
> Atenciosamente,

A assinatura configurada no Outlook é adicionada automaticamente ao final da mensagem.

---

# 📎 Anexo automático

O arquivo gerado durante a etapa do Excel é anexado automaticamente ao e-mail.

Exemplo:

```
MAPA DE VENDAS.pdf
```

---

# ⚙️ Como executar

Instale as dependências:

```bash
pip install xlwings pywin32
```

Depois execute:

```bash
python main.py
```

Ao término da geração do PDF:

```bash
python envio_email.py
```

Caso os dois scripts sejam integrados em um único fluxo, basta executar o arquivo principal.

---

# 📋 Requisitos

* Windows
* Microsoft Excel
* Microsoft Outlook Desktop
* Python 3.10+
* xlwings
* pywin32
* Arquivo Excel configurado
* Conexões do Power Query funcionando

---

# 💡 Melhorias futuras

* Remover o uso de `time.sleep()` utilizando espera inteligente.
* Modularizar o projeto em funções.
* Adicionar tratamento de exceções.
* Gerar arquivos de log.
* Configurar caminhos através de arquivo `.env`.
* Permitir configuração dinâmica dos destinatários.
* Enviar o e-mail automaticamente utilizando `mail.Send()`.
* Agendar a execução pelo Agendador de Tarefas do Windows.
* Criar interface gráfica para execução da automação.

---

# 👨‍💻 Autor

**Murilo Moraes**

Projeto desenvolvido para automatizar o processo diário de geração e distribuição do Mapa de Vendas, reduzindo atividades manuais, aumentando a confiabilidade dos dados e agilizando o envio das informações para toda a equipe.
