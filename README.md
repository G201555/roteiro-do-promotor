# Roteiro do promotor

Transforma o roteiro mensal (uma planilha grande com todas as lojas e dias) em entregáveis prontos para a equipe e para o sistema de campo.

## Scripts

| Arquivo | O que faz |
|---|---|
| `gerar.py` | Gera o roteiro de um promotor ou de todos, com resumo, roteiro e agenda por dia |
| `upload.py` | Gera a planilha de importação do sistema: uma linha por loja e semana ISO, com os dias marcados |
| `validar.py` | Confere o resultado: lojas sem rota, frequência diferente dos dias marcados, visitas por dia |

## Stack

Python, openpyxl.

---
Projeto feito no trabalho, publicado sem dados.
