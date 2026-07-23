# KPI - Mapa de Vendas (automação xlwings)

Script em Python que automatiza a atualização e o ajuste do relatório "Mapa de Vendas", eliminando o trabalho manual de filtrar e copiar linhas com valores inconsistentes (`#N/D`, `0`, vazio, `#VALOR!`) para uma aba de ajuste.

> Este projeto está em desenvolvimento e foi construído como estudo prático de xlwings, então o código e este README evoluem junto com o aprendizado.

## O que o script faz

1. Abre o Excel via `xlwings` e carrega a planilha `Mapa de vendas`.
2. Ativa a aba `Pedido de venda` e força a atualização de todas as conexões de dados (`RefreshAll`).
3. Descobre dinamicamente o tamanho real da tabela de pedidos (última linha e última coluna com dado).
4. Aplica um filtro (`AutoFilter`) na coluna `Canal de venda oficial`, mantendo visível apenas as linhas com `#N/D`.
5. Copia só as linhas visíveis (`SpecialCells` tipo 12 — visível) e cola no final dos dados já existentes na aba `Filtro - canal ajustado`.
6. Na aba de destino, aplica um segundo filtro por múltiplos critérios (`#N/D`, `0`, vazio, `#VALOR!`) para isolar o que precisa de ajuste manual.

## Como rodar

```bash
python main.py
```

Requer o arquivo Excel de origem acessível no caminho configurado no script (rede corporativa) e o Excel instalado na máquina (o xlwings controla uma instância real do Excel).

## Estrutura conceitual do fluxo

```
Abrir arquivo (xlwings)
  → Aba "Pedido de venda"
      → RefreshAll() + espera
      → Descobrir range da tabela (última linha/coluna)
      → Filtrar coluna P por "#N/D"
      → Copiar só linhas visíveis
  → Aba "Filtro - canal ajustado"
      → Colar no final dos dados existentes
      → Filtrar coluna N por múltiplos critérios de inconsistência
```

## Conceitos-chave usados (para quem está aprendendo, como eu)

- **`.end('down')` / `.end('right')`**: simula Ctrl+Seta no teclado para achar dinamicamente onde os dados terminam, sem precisar fixar números de linha/coluna no código.
- **`.api`**: ponte do xlwings para os métodos nativos do Excel/VBA (usado para `AutoFilter` e `SpecialCells`, que não têm atalho direto no xlwings).
- **`AutoFilter` não retorna a tabela filtrada** — ele só ativa o filtro (esconde linhas). Quem devolve as células realmente visíveis é `SpecialCells(12)`, chamado *depois* do filtro, sobre o mesmo range.
- **Filtro por múltiplos valores**: `Criteria1` como lista + `Operator=7` (`xlFilterValues`), já que o `or` do Python não funciona como "múltiplas opções" dentro de uma string.

## Problema conhecido / em investigação

`RefreshAll()` às vezes não atualiza a conexão de dados mesmo quando o clique manual em "Atualizar Tudo" também falha — suspeita é a propriedade de conexão "Atualizar esta conexão ao usar Atualizar Tudo" estar desmarcada nas Propriedades da Conexão (aba Dados → Conexões).

## Melhorias futuras planejadas

- Substituir `time.sleep(15)` por uma espera real do término do `RefreshAll()`.
- Tratar o caso da aba de destino estar completamente vazia (primeira execução), onde `.end('down')` se comporta de forma diferente.
- Modularizar em arquivos separados (abertura do Excel, filtro, cópia/colagem) em vez de um único script.