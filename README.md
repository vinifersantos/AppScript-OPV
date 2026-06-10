# App Script — DRE Grupo OPV

Script unificado para Google Sheets que automatiza o controle financeiro do Grupo OPV.  
Contém dois módulos independentes: **Lançamentos** e **Boletos Pagos**.

---

## Módulos

### 1. Lançamentos

Registra receitas e despesas na aba `LANÇAMENTOS` com validação completa de dados.

**Menu gerado:** `Lançamentos`

| Item de menu | Função |
|---|---|
| Registrar lançamento | Valida e grava o formulário na tabela |
| Excluir lançamentos marcados | Remove linhas com checkbox marcado |
| Adicionar item às listas | Cadastra novo Tipo + Categoria + Subcategoria na aba `CONFIG` |
| Limpar formulário | Apaga os campos do formulário |

**Abas necessárias:**

- `LANÇAMENTOS` — formulário em `A4:C14`, cabeçalho na linha 19 e dados nas linhas 20–5000
- `CONFIG` — lista de Tipo / Categoria / Subcategoria / Classe no relatório

**Intervalos nomeados obrigatórios:**

| Nome | Descrição |
|---|---|
| `frmDataLancamento` | Campo Data |
| `frmMesCompetencia` | Campo Mês de competência (opcional) |
| `frmUnidade` | Campo Unidade |
| `frmTipoLancamento` | Campo Tipo |
| `frmCategoria` | Campo Categoria |
| `frmSubcategoria` | Campo Subcategoria |
| `frmValor` | Campo Valor |
| `frmFormaPgto` | Campo Forma de Pagamento |
| `frmStatus` | Campo Status |
| `frmDataPgto` | Campo Data de Pagamento |
| `frmObservacoes` | Campo Observações |

O script prioriza o mapa fixo atual do formulário (`B4:B14`) e usa intervalos nomeados apenas como compatibilidade quando necessário.

**Colunas da tabela `LANÇAMENTOS`:**

| Col | Conteúdo |
|---|---|
| A | ID |
| B | Data |
| C | Mês |
| E | Unidade |
| F | Tipo |
| G | Categoria |
| H | Subcategoria |
| I | Valor |
| J | Forma de Pagamento |
| K | Status |
| L | Data de Pagamento |
| M | Observações |
| N | Criado em |
| O | Excluir (checkbox) |
| P | Chave Técnica (anti-duplicidade) |
| CL | Classe do tipo (fórmula/automação) |
| R | Lista de Tipos |
| S | Lista de Categorias |
| T | Lista de Status |
| U | Lista de Formas de Pagamento |

**Classes de relatório disponíveis no fluxo operacional de novos itens:**
`Despesa` · `Ignorar`

Ao usar **Lançamentos → Adicionar item às listas**, o item é salvo na `CONFIG` sempre com `ClasseRelatorio = Despesa` e o script pergunta se deve refletir a nova linha nas abas de lojas/DRE:

- `CENTRO`
- `C. DO MAR`
- `OUTLET`
- `TNC.`

A inserção nas lojas é conservadora: o script procura uma linha-modelo por mesma Categoria ou ClasseRelatorio, copia formatação/fórmulas/validações/altura e altera apenas campos textuais identificados. Se não encontrar uma estrutura segura, não insere automaticamente e mostra pendência para revisão manual.

---

### 2. Boletos Pagos

Registra boletos e controla pagamentos na aba `BOLETOS PAGOS` com fórmulas automáticas de status, semana e diferença de valores.

**Menu gerado:** `Boletos`

| Item de menu | Função |
|---|---|
| Registrar boleto | Valida e grava o formulário na tabela |
| Limpar formulário | Apaga os campos do formulário |
| Configurar / revisar tudo | Executa setup completo (cabeçalhos, validações, fórmulas, layout) |
| Configurar validações | Reaplica dropdowns com base na aba `CONFI` |
| Reaplicar fórmulas automáticas | Recria todas as fórmulas nas colunas automáticas |
| Adicionar item às listas | Adiciona Loja, Empresa, Categoria ou Forma de Pagamento na `CONFI` |
| Verificar estrutura da planilha | Diagnóstico completo da estrutura esperada |

**Abas necessárias:**

- `BOLETOS PAGOS` — tabela de dados (linhas 15–5000)
- `CONFI` — listas de Loja, Empresa, Categoria, Forma de Pagamento, Status
- `RELATÓRIO DE PAGAMENTOS` — aba de relatório (exigida pelo setup)

**Formulário — intervalo `B4:B12`:**

| Linha | Campo |
|---|---|
| B4 | Data de Vencimento |
| B5 | Loja |
| B6 | Empresa / Fornecedor |
| B7 | Descrição / Observação |
| B8 | Valor a Pagar |
| B9 | Valor Pago |
| B10 | Forma de Pagamento |
| B11 | Data de Pagamento |
| B12 | Observações |

**Colunas automáticas (geradas por fórmula `ARRAYFORMULA`):**

| Col | Conteúdo |
|---|---|
| B | Dia da Semana |
| C | Semana do Mês |
| H | Status (espelho de R) |
| K | Diferença (previsto − pago) |
| N | Mês (texto, maiúsculo) |
| O | Nº do Mês |
| P | Ano |
| Q | Semana do Mês (relatório) |
| R | Status Automático |
| S | Diferença Automática |
| T | Valor em Aberto |
| U | Valor em Atraso |
| V | Quantidade |

**Colunas manuais (preenchidas pelo script):**

| Col | Conteúdo |
|---|---|
| A | Data de Vencimento |
| D | Loja |
| E | Empresa |
| F | Descrição |
| G | Forma de Pagamento |
| I | Valor Previsto |
| J | Valor Pago |
| L | Data de Pagamento |
| M | Observações |

---

## Como instalar

1. Abra o Google Sheets da planilha do Grupo OPV.
2. Acesse **Extensões → Apps Script**.
3. Apague qualquer conteúdo existente e cole o conteúdo de `appscript.gs`.
4. Salve (`Ctrl + S`) e recarregue a planilha.
5. O menu **Lançamentos** e o menu **Boletos** aparecerão na barra superior.
6. No módulo Boletos, execute **Boletos → Configurar / revisar tudo** para aplicar cabeçalhos, validações e fórmulas na primeira utilização.

---

## Estrutura do repositório

```
appscript.gs   — código-fonte completo (único arquivo a colar no Apps Script)
README.md      — esta documentação
```

---

## Regras gerais do código

- Existe apenas um `onOpen()` no projeto; ele chama os dois menus.
- Todas as funções do módulo Lançamentos têm prefixo `LA_`.
- Todas as funções do módulo Boletos têm prefixo `BP_`.
- Helpers compartilhados têm prefixo `APP_`.
- Lock de documento (`LockService`) é usado nas operações de escrita para evitar conflitos em edições simultâneas.
- Cache de documento (`CacheService`, TTL 5 min) é usado para a leitura das listas da aba `CONFI`.
