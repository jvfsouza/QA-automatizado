# Prompt para gerar a planilha de perguntas com uma IA

Prompt reutilizável para pedir a uma IA (ChatGPT, Claude, Gemini etc.) uma planilha de perguntas no formato que o bot lê. As regras do formato estão em `FORMATO_PLANILHA.md`.

**Como usar:** copie o bloco abaixo, preencha o que está entre `[colchetes]` e cole na IA. Depois siga a seção "Depois que a IA responder".

---

## Prompt

```text
Você vai criar casos de teste para o chat de atendimento de um site. Quero o resultado em um formato que vira uma planilha .xlsx.

CONTEXTO
- Site: [nome do site e o que ele faz, ex.: casa de apostas online / loja de eletrônicos]
- Quem responde no chat: [ex.: assistente de IA do site]
- Assuntos que o atendimento deveria saber: [ex.: cadastro, verificação de conta, depósito, saque, bônus, senha, jogo responsável]
- Idioma: português do Brasil, com linguagem natural de cliente (frases curtas, às vezes informais)

O QUE CRIAR
- [10] conversas. Cada conversa tem de 1 a [4] perguntas, enviadas em ordem na mesma conversa.
- Em conversas com mais de uma pergunta, as seguintes devem depender do contexto das anteriores (ex.: "Como faço um saque?" → "Qual o prazo?" → "E se for no fim de semana?").
- Misture estes tipos de caso:
  1. perguntas comuns sobre os assuntos acima;
  2. conversa encadeada que depende do contexto;
  3. pergunta fora do escopo do site;
  4. texto sem sentido seguido de uma correção ("asdkjh qwerty" → "desculpe, quero saber sobre X");
  5. pedido de atendente humano;
  6. pergunta longa e detalhada;
  7. pergunta sensível ou de risco (ex.: menor de idade, bloquear a conta, reclamação);
  8. saudação e despedida.
- Não repita perguntas. Não invente nomes de produtos, valores ou regras do site como se fossem fatos: as perguntas devem ser apenas o que um cliente perguntaria.

FORMATO DA RESPOSTA
Responda com UM ÚNICO bloco de código CSV, separado por ponto e vírgula, sem nenhum texto antes ou depois do bloco:
- A primeira linha é o cabeçalho: Pergunta 1;Pergunta 2;...;Pergunta N, onde N é o maior número de perguntas de uma conversa.
- Cada linha seguinte é uma conversa. Deixe vazias as colunas que sobrarem (ex.: uma conversa de 2 perguntas em uma tabela de 4 colunas termina com ;;).
- Não use cabeçalhos diferentes de "Pergunta 1", "Pergunta 2" etc. Não adicione colunas extras.
- Envolva em aspas duplas qualquer pergunta que contenha ponto e vírgula ou aspas.
- Nenhuma pergunta pode começar com =, +, - ou @.
```

---

## Depois que a IA responder

1. **Revise as perguntas.** Remova as que o site claramente não deveria responder, se não forem casos propositais de "fora do escopo".
2. **Converta para `.xlsx`:**
   - Salve o CSV como `perguntas.csv` em UTF-8 e abra no Excel/LibreOffice escolhendo `;` como separador; depois salve como `.xlsx`, ou
   - rode o script abaixo, que faz a conversão.
3. Salve em `venv/perguntas.xlsx` (veja `FORMATO_PLANILHA.md` para as regras) e rode um teste curto antes do completo.

### Script de conversão CSV → xlsx

```python
import csv
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
with open("perguntas.csv", encoding="utf-8-sig", newline="") as f:
    for linha in csv.reader(f, delimiter=";"):
        ws.append([c.strip() or None for c in linha])
wb.save("venv/perguntas.xlsx")
```

Rode com `.venv/bin/python` na raiz do projeto, com `perguntas.csv` na mesma pasta.

## Variações do prompt

- **Mais cobertura de um assunto:** troque a lista de tipos de caso por "todas as conversas devem ser sobre [assunto], com variações de jeito de perguntar".
- **Teste de regressão:** cole as perguntas de uma planilha anterior e peça "reescreva cada pergunta de outro jeito, mantendo a intenção".
- **Casos de erro do cliente:** peça erros de digitação, gírias, perguntas sem acento, tudo em maiúsculas e mensagens muito curtas ("saque").
- **Teste de limites:** peça perguntas com 300+ caracteres, perguntas com vários assuntos de uma vez e a mesma pergunta repetida duas vezes na mesma conversa.
