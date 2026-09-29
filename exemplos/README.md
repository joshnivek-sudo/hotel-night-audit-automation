# Exemplos (dados fictícios)

Tudo nesta pasta é **inventado**. Nenhum número, nome, quarto ou valor veio de um hotel real.

Os relatórios imitam só a **estrutura** que os scripts leem — onde fica cada rótulo e cada número. Os modelos de planilha (cartões, Estatístico e Folha de Rosto) foram **desenhados do zero** pra demonstração: não são cópia das planilhas usadas em nenhum hotel.

## O que tem aqui

| Arquivo | O que é |
|---|---|
| `cartoes/RELACAO_DE_COMANDAS_DEMO.xlsx` | Relatório de comandas com 22 lançamentos (maquininha, online, PIX, erro de digitação e uma bandeira desconhecida) |
| `cartoes/MODELO_CONTROLE_DE_CARTOES.xlsx` | Modelo da planilha de cartões |
| `estatistico/RDS_RELATORIO_DEMO.xlsx` | Resumo Diário de Situação de 28/09/2026 |
| `estatistico/on_the_book_demo.xlsx` | Reservas dos 14 dias seguintes |
| `estatistico/conta_pendente_demo.xlsx` | 7 contas em aberto |
| `estatistico/situacao_uhs_demo.xlsx` | 40 UHs do hotel + 8 de condomínio, com walk-ins, layovers e um mensalista |
| `estatistico/ESTATISTICO_HOTEL_DEMO.xlsx` | Estatístico com o mês em andamento (dias 1 a 27 já lançados) |
| `estatistico/FOLHA_DE_ROSTO_DEMO.xlsx` | Modelo da folha de rosto |
| `gerar_exemplos.py` | Script que gera todos os arquivos acima |

## Rodando a demonstração

**Cartões**

1. Copie os dois arquivos de `exemplos/cartoes/` para a pasta `cartoes/`
2. Dê dois cliques em `cartoes/EXECUTAR.bat` (ou rode `python preencher_cartoes.py` dentro da pasta)
3. A planilha pronta aparece em `cartoes/PRONTAS/`

**Estatístico**

1. Copie os seis arquivos de `exemplos/estatistico/` para a pasta `estatistico/`
2. Dê dois cliques em `estatistico/Preencher Estatistico.bat` (em Linux ou Mac: `python "Preencher Estatistico.bat"`)
3. Os arquivos saem com a data no nome: `..._28-09-2026.xlsx`. O resultado esperado termina com `BATEU CERTINHO`

> Copie em vez de rodar direto aqui: o Estatístico renomeia a planilha e manda a versão anterior pro backup. Dentro de `cartoes/` e `estatistico/`, as planilhas geradas são ignoradas pelo `.gitignore` e não vão pro GitHub por engano.

## Gerando de novo

```bash
python exemplos/gerar_exemplos.py
```

Os números são aleatórios com semente fixa, então saem sempre iguais — e os testes em `tests/` contam com isso.
