# Relatório Financeiro — Doce Sabor

Dashboard mobile-first, em pt-BR, do controle financeiro do restaurante, lido da planilha
**Controle Doce Sabor.xlsx** (Excel Online).

## Arquivos

| Arquivo | Papel |
|---|---|
| `index.html` / `styles.css` / `app.js` | o site (sem dependências externas) |
| `data.json` | dados agregados — gerado |
| `aggregate.ts` | lê a planilha via Graph e agrega (função `buildReportData`) |
| `generate-report.ts` | CLI: `npm run report` regrava `data.json` |
| `server.ts` | servidor: serve o site e mantém o `data.json` atualizado |

## O que mostra

- **Resultado do mês** + 4 cartões (Entradas, Saídas, Despesas fixas, Saldo).
- **Gastos por categoria**, **formas de pagamento** (receitas e despesas), **fornecedores**,
  **despesas fixas** e **navegação por mês**.

## Dois modos de rodar

O navegador **não consegue ler a planilha** (o Microsoft Graph bloqueia CORS e o token não pode
ir para o dispositivo do cliente). Então há duas formas:

1. **Estático** — você gera o `data.json` na sua máquina e publica só os arquivos. O botão "🔄"
   apenas recarrega o `data.json` já publicado.
2. **Ao vivo** — um servidor Node guarda o token, lê a planilha e serve o `data.json` fresco
   (no intervalo definido e sob demanda). É o único modo em que o botão "🔄" puxa dados novos.

## Rodar localmente

```bash
cd /Volumes/Storage/www/doce-sabor/excel-mcp

npm run report            # modo estático: regrava report/data.json
npm run report:serve      # modo ao vivo (TS): http://localhost:8787
npm run build && npm run report:start   # modo ao vivo (compilado): node dist/report/server.js
```

| Env | Padrão | O que faz |
|---|---|---|
| `PORT` | `8787` | porta |
| `REFRESH_MINUTES` | `60` | intervalo entre atualizações automáticas |
| `REFRESH_TOKEN` | *(vazio)* | senha opcional exigida em `POST /refresh?token=…` |
| `REPORT_ROOT` | `report/` do projeto | pasta dos arquivos estáticos (não precisa mudar) |

Endpoints do servidor:

- `GET /` e arquivos estáticos — o relatório.
- `GET /data.json` — os dados (atualiza se estiver desatualizado).
- `POST /refresh` — lê a planilha na hora (é o que o botão "🔄" chama).

## Deploy no Coolify (modo ao vivo — botão funciona)

Crie um recurso **Application** (Node), não um site estático.

- **Build command**: `npm ci && npm run build`
- **Start command**: `npm run report:start`   (ou `node dist/report/server.js`)
- **Port**: `8787`

**Variáveis de ambiente** (coloque no painel do Coolify — funcionam como secrets):

| Variável | Valor |
|---|---|
| `PORT` | `8787` |
| `REFRESH_MINUTES` | `15` |
| `EXCEL_SOURCE` | `graph_download` |
| `EXCEL_DRIVE_ID` | `62173A206D2F8FA1` |
| `EXCEL_ITEM_ID` | `62173A206D2F8FA1!s38b60b88feed4d51856f4cfacba51b0c` |
| `EXCEL_SCOPES` | `offline_access Files.Read.All` |
| `EXCEL_REFRESH_TOKEN` | *(veja abaixo)* |

**Pegando o `EXCEL_REFRESH_TOKEN`** (na sua máquina, onde você fez o login):

```bash
cat .tokens/graph.json | python3 -c "import sys,json;print(json.load(sys.stdin)['refresh_token'])"
```

Cole a saída como valor de `EXCEL_REFRESH_TOKEN`. É o segredo que deixa o servidor renovar o
acesso à planilha — não coloque em repositório nem compartilhe.

**Segurança do endpoint** (recomendado): defina também `REFRESH_TOKEN` e, no `app.js`, troque
`fetch("refresh", …)` por `fetch("refresh?token=SEU_TOKEN", …)` para o `POST /refresh` não ficar
aberto ao público.

> **Por que `graph_download`?** A planilha está numa conta **pessoal** (onedrive.live.com) e a
> API de workbook do Graph não suporta contas pessoais. `graph_download` resolve o arquivo pelo
> Graph e baixa o `.xlsx`, que é o caminho suportado nesse caso.

## Deploy no Coolify (modo estático — sem botão)

Se preferir não manter servidor: publique só os arquivos de `report/` como site estático e, para
atualizar, rode `npm run report` na sua máquina e faça deploy de novo (o botão vira só um recarregar).

## Notas

- `data.json` contém totais agregados (não as linhas individuais) — o cliente vê exatamente o que
  o relatório mostra.
- Se os números parecerem errados, confira se a aba `Entradas` está preenchida (o valor atual é só
  de exemplo) e se a coluna `Categoria` está consistente (o relatório normaliza "frios"/"Frios").
- As abas lidas são `Lançamentos`, `Entradas` e `Contas fixas`. `Consultas`/`Cadastros`/`Menu` não
  entram no relatório.
