# Financa - Sistema Desktop Financeiro

Aplicativo desktop local para organizacao financeira pessoal, com foco em:

- fluxo de caixa empresarial e pessoal;
- controle de contas e gastos;
- acompanhamento de investimentos;
- dados locais em SQLite;
- backup automatico por alteracao;
- interface estilo planilha com AG Grid.

## Tecnologias

- Electron
- Node.js
- SQLite (`better-sqlite3`)
- Tailwind CSS
- AG Grid

## Como rodar localmente

```bash
npm install
npm start
```

## Build do instalador Windows (.exe)

```bash
npm run build:win
```

## Download para usuarios finais

- Release mais recente: [Baixar no GitHub Releases](https://github.com/Thiago-DD/Financa/releases/latest)
- Link direto esperado do instalador: `https://github.com/Thiago-DD/Financa/releases/latest/download/FinanceiroPessoal-Setup-1.0.0.exe`

## Estrutura principal

- `main.js`: processo principal do Electron, IPC, polling de cotacoes e notificacoes.
- `database.js`: criacao de tabelas, seed e funcoes de persistencia.
- `preload.js`: bridge segura entre front-end e Electron.
- `index.html`: layout e tema.
- `renderer.js`: logica da interface, grids e atualizacoes reativas.
