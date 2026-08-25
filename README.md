# Formulário ADM CTI

Formulário de linha de comando em Python para coleta estruturada de dados 
de admissão de pacientes em CTI (Centro de Terapia Intensiva).

## Funcionalidades
- Cálculo automático de idade a partir da data de nascimento
- Classificação do nível de assistência com base na escala IMS (0-10)
- Registro condicional de dados de VMI (ventilação mecânica invasiva)
- Registro condicional de oxigenoterapia (fluxo em L/min)
- Registro de histórico de tabagismo (ativo, cessado ou não fumante)
- Exportação incremental para Excel (.xlsx), preservando o histórico 
  de atendimentos anteriores
- Dados prontos para conexão e análise no Power BI

## Tecnologias
- Python
- openpyxl (manipulação de arquivos Excel)
