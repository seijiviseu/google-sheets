# Código para automação do Google Sheets do Rodrigo

Este código foi feito em python usando a biblioteca *gspread*.
Os dados são coletados por fora via API do Conta Azul

## Fluxo

1. Declaração das credenciais de acesso, tokens e API do Google
2. Coleta dos dados já presentes na planilha para validação
3. Processamento dos dados da API do Conta Azul
4. Filtragem dos dados que já existem e do mês
5. Ordenação dos nomes dos vendedores por ordem alfabética
6. Envio dos dados filtrados ao Google Sheets
7. Formatação das células conforme o status das parcelas