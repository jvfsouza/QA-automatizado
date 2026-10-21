# Formato da planilha de perguntas

Como criar o `venv/perguntas.xlsx` que o bot lê. Para gerar a planilha com ajuda de uma IA, use o prompt de `PROMPT_GERAR_PLANILHA.md`. O arquivo `venv-example/perguntas.xlsx` é um exemplo pronto.

## Modelo

| Pergunta 1 | Pergunta 2 | Pergunta 3 | Pergunta 4 |
|---|---|---|---|
| Olá, bom dia! | Como funciona o atendimento por aqui? | | |
| Como faço um saque? | Qual o prazo para o dinheiro chegar? | Existe valor mínimo para sacar? | |
| Preciso falar com um atendente humano | | | |

- **Cada linha é uma conversa.** O bot abre uma sessão nova do chat por linha.
- **Cada coluna é uma pergunta**, enviada na ordem das colunas, na mesma conversa. As perguntas seguintes podem depender das anteriores (ex.: "Quero saber meu pedido" → "O número é 123456").
- Uma conversa pode ter 1 pergunta ou várias. Não há limite de colunas.

## Regras

| Regra | Detalhe |
|---|---|
| Cabeçalho na linha 1 | Colunas cujo nome **começa com "Pergunta"** (sem diferenciar maiúsculas) são lidas. Qualquer outra coluna é ignorada, então dá para ter colunas de anotação, como "Cenário" ou "Observação" |
| Número da pergunta | Vem do número no cabeçalho (`Pergunta 3` → `pergunta_3.png`), não da posição. Sem número no nome, vale a ordem entre as colunas de pergunta |
| Células vazias | São ignoradas, inclusive no meio da linha. As perguntas restantes continuam em ordem |
| Linhas sem nenhuma pergunta | São puladas |
| Aba | Só a **aba ativa** (a que estava aberta ao salvar) é lida. Mantenha uma única aba, ou deixe a das perguntas ativa |
| Formato do arquivo | `.xlsx`. Não use `.xls`, `.csv` ou `.ods` |
| Nome e local | `venv/perguntas.xlsx` (definido por `ENTRADA` no topo do `bot_teste.py`) |
| Texto | Cada pergunta é lida como texto, com espaços nas pontas removidos. Números e datas viram texto |

A numeração das linhas no resultado conta o cabeçalho como linha 1: a primeira conversa é a **linha 2** (`output/prints/linha_2/...`), igual ao que o Excel mostra.

## Como criar

### No Excel ou LibreOffice

1. Crie uma planilha nova.
2. Na linha 1, escreva `Pergunta 1`, `Pergunta 2`, ... até o maior número de perguntas de uma conversa.
3. Escreva uma conversa por linha, uma pergunta por célula.
4. Salve como `venv/perguntas.xlsx` (formato Excel `.xlsx`).

### Com Python (openpyxl)

```python
from openpyxl import Workbook

conversas = [
    ["Olá, bom dia!", "Como funciona o atendimento por aqui?"],
    ["Como faço um saque?", "Qual o prazo para o dinheiro chegar?", "Existe valor mínimo para sacar?"],
    ["Preciso falar com um atendente humano"],
]

wb = Workbook()
ws = wb.active
ws.append([f"Pergunta {i}" for i in range(1, max(map(len, conversas)) + 1)])
for c in conversas:
    ws.append(c)
wb.save("venv/perguntas.xlsx")
```

## Cuidados

- **Texto que começa com `=`, `+`, `-` ou `@`** pode virar fórmula no Excel e a célula perde o texto original. Se precisar mesmo, escreva com um apóstrofo na frente (`'=1+1`) ou formate a coluna como texto antes de digitar. Em Python, use `cell.data_type = "s"` depois de gravar.
- **Salvar pelo LibreOffice** um arquivo que tinha texto literal começando com `=` pode reconvertê-lo em fórmula.
- **Feche o arquivo** antes de rodar o bot, ou ao menos não o salve durante a execução.
- **Perguntas realistas:** o bot do site só deve ser testado com o que ele deveria saber responder. Misture casos felizes, perguntas fora do escopo, texto sem sentido, pedido de atendente humano e conversas encadeadas.
- **Poucas perguntas por conversa** deixam a execução mais rápida: cada resposta leva cerca de 20 s no Brazino777.

## Erros comuns

| Sintoma | Causa |
|---|---|
| `Nenhuma coluna 'Pergunta N' encontrada` | O cabeçalho não está na linha 1 ou nenhuma coluna começa com "Pergunta" |
| O bot lê 0 conversas e termina | Todas as células de pergunta estão vazias ou a aba ativa é outra |
| `FileNotFoundError` | O arquivo não está em `venv/perguntas.xlsx` |
| Pergunta virou número ou data | A célula foi formatada como número/data; formate a coluna como texto |
