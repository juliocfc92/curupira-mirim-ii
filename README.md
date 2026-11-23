# Operação Curupira-Mirim II — Registro de Ações

Sistema de coleta de dados da **Operação Integrada Curupira-Mirim II** (2ª Ed. Amazonas) —
combate à violência contra crianças e adolescentes indígenas na tríplice fronteira do Alto Solimões
(Tabatinga, Benjamin Constant e Atalaia do Norte).

**Coordenação:** SENASP/MJSP · Polícia Civil do Amazonas (PC-AM)

## Arquivos
- `index.html` — formulário de registro + painel de indicadores (página única, sem dependências)
- `curupira-mirim-gas.js` — backend Google Apps Script (grava na planilha de respostas)

## Indicadores (Quadro 3 do PAI)
Abrangência · Ações Preventivas e Educativas · Diligências · Ações Repressivas · Ações Cartorárias

## Deploy
Página estática — pode ser servida pelo GitHub Pages. O backend roda no Google Apps Script
(a URL `/exec` fica embutida em `index.html`, variável `GAS_URL`).
