# * APN Fuel Sales ETL Challenge 2021

O desafio tem como finalidade extrair dados de tabelas dinamicas de planilhas em formato xls e criar um pipeline de dados que atenda aos seguintes requisitos:
- Vendas de combustiveis derivados de óleo por estado(UF) e produto;
- Vendas de Diesel por estado(UF) e tipo.

Premissas:
- O total de dados extraidos deve manter a quantidade do total das tabelas dinamicas.
- A formatação dos dados deve seguir o seguinte conceito:

Column Type
year_month date
uf string
product string
unit string
volume double
created_at timestamp


* Use Case / Challenges
Devido a origem dos dados seguir o conceito de pivot tables (tabelas dinamicas) isso se mostrou um grande desafio logo no início. 

A abordagem adotada foi:

- utilizando o módulo do python (win32com) foi feita a conversão da tabela de *.xls para *.xlsx;
- em seguida uma lógica para transformar o arquivo *.xlsx em *.xml

Obs: Não foi possível encontrar uma solução compatível ou algo similar ao win32 para arquitetura Linux forçando o uso de um sistema operacional Windows para o passo de conversão de extensão *.xls para *.xlsx.
     (Caso esse passo possa ser executado de forma manual o código completo se torna compatível para arquiteturas Linux).

- Com os objetos xml a disposição foi necessário o entendimento da interdepência entre os objetos "pivotCacheDefinition" e "pivotCacheTables".
- Com o uso da biblioteca pandas foi possível atender a necessidade do desafio utilizando a criação de dataframes e organizando os dados em schemas.