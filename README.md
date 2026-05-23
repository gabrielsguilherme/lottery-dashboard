# lottery-dashboard

Aplicação web local em Node.js para analisar resultados da Mega-Sena a partir da planilha Excel, sem usar SQLite.

## Stack
- Node.js
- Express
- ExcelJS
- Chart.js
- HTML/CSS/JavaScript

## Como executar no Mac

1. Instale as dependências:
   ```bash
   npm install
   ```
2. Coloque sua planilha em:
   ```bash
   data/MegaSena_Dashboard.xlsx
   ```
3. Rode o projeto:
   ```bash
   npm start
   ```
4. Abra no navegador:
   ```
   http://localhost:3000
   ```

## Funcionalidades
- Leitura direta da planilha Excel
- Últimos resultados
- Frequência histórica das dezenas
- Estatísticas de soma e paridade
- Gerador de jogos com filtros
- Recarregamento da planilha sem reiniciar o app

## Estrutura
```
lottery-dashboard/
├── data/
│   └── MegaSena_Dashboard.xlsx
├── public/
│   ├── index.html
│   └── app.js
├── package.json
├── server.js
└── README.md
```
