# 🎰 Lottery Dashboard

Dashboard web local para análise e visualização de resultados de loteria.

## Tecnologias

- **Node.js + Express** — servidor local
- **Chart.js** — gráficos interativos
- **XLSX** — leitura de planilhas Excel e CSV

## Instalação

```bash
# 1. Instale as dependências
npm install

# 2. Inicie o servidor
npm start

# 3. Acesse no navegador
open http://localhost:3000
```

## Desenvolvimento (com hot-reload)

```bash
npm run dev
```

## Funcionalidades

| Seção | Descrição |
|---|---|
| **Visão Geral** | KPIs e gráficos resumidos (top 10, distribuição por faixas) |
| **Frequência** | Frequência de cada número, ordenável, com grid colorido |
| **Histórico** | Tabela paginada com busca de todos os sorteios |
| **Análise** | Números atrasados, paridade (pares/ímpares), heatmap de temperatura |
| **Importar** | Upload de planilha (.xlsx, .xls, .csv) via drag & drop |

## Formato da Planilha

A planilha pode ter qualquer estrutura. O sistema detecta automaticamente colunas numéricas com valores entre 1 e 100 como dezenas sorteadas. Colunas sugeridas:

```
Concurso | Data | Bola1 | Bola2 | Bola3 | Bola4 | Bola5 | Bola6
```

## Licença

MIT
