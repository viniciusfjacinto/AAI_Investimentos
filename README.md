# Agente Autônomo de Investimentos

Projeto desenvolvido para gerenciar a carteira de Clientes de uma empresa de investimentos. O dashboard completo anonimizado pode ser visto em: https://app.powerbi.com/view?r=eyJrIjoiNGZhMThjNWUtMTRmMS00M2I3LWI3NTYtYTg3ZDgyYzFjMjliIiwidCI6ImM3Mzk2ZTVlLTYzMDYtNGIwZi1hN2NmLWI1YzFhNDRkNDk0MSJ9

A estrutura do projeto foi dividida em 3 etapas:

# 1 - Preparação da Base de Dados em Excel

Obs: Foi escolhido o Excel devido à necessidade de usufruto dos dados da plataforma Economática. Os quais, de acordo com o plano assinado pela empresa, só são possíveis de serem extraídos utilizando um Add-In disponível na ferramenta da Microsoft.

Na base em Excel dividmos os dados em planilhas diferentes para facilitar a identificação e conexão destes dados em outras ferramentas. Assim, encontramos as seguintes planilhas:

#### Fundos - Informações generalizadas sobre os fundos utilizados pela empresa
  *Obs: Possui 2 linhas com nomes de colunas para facilitar a fixação de nomes no tratamento pelo Power BI, dado que os nomes das colunas na linha 2 se alteram diariamente sempre que atualizadas

#### Retorno Fundos 1Ydaybyday - Retornos dos fundos no período entre o último dia útil (D-1) e o mesmo dia no ano passado. A contagem segue apenas os dias úteis

#### Cotas - Histórico do valor da cota dos fundos no período de 5 anos

#### IBOV CDI - Valores históricos de fechamento do Ibovespa e o CDI Acumulado no período de 5 anos

#### Base Molde - Planilha onde são colados os valores da tabela de Operações (extraída do BI do Safra Invest) para obtermos os valores de Retorno,  Cota e Volatilidade a partir das data de aplicação/resgate em cada Fundo para cada Cliente, após obedecidos os seguintes critérios:
  - Remove as colunas não essenciais, ficando apenas com Data, Cliente, Mercadoria/Fundo, Quantidade e Valor
  - Filtro em Mercadoria/Fundo remove COE, Debêntures e VIS
  - Filtro em Valores remove valores entre -100 e 0
  - Aplicação de PROCV para adicionar uma coluna com CNPJs e substituir a coluna Cliente pelos nomes dos Clientes (por meio das planilhas dFundos e dClientes)
*Obs: Os dados são colados na forma vertical/wide (onde cada coluna representa uma movimentação) pois isso acelera em muito o tempo de realização das consultas pela Economática, sendo a forma mais eficiente para que as informações possam ser atualizadas várias vezes ao dia

#### Base Safra - Base de dados onde são colados os valores da planilha 'Base Molde', transpostos e abaixo do cabeçalho

#### dFundos - Relação entre Fundos e CNPJ de acordo com os nomes distintos presentes na planilha de Operações, coluna Mercadoria/Fundos

#### dClientes - Relação entre código do Cliente (presente na coluna Cliente da planilha de Operações), nome e assessor responsável, proveniente da tabela Rateio Fundo Assessores

*Também há uma planilha oculta chamada Apoio, onde foram colados valores na horizontal para facilitar as fórmulas para cálculo do retorno na planilha Fundos 1Ydaybyday

Cuidados Essenciais: A planilha dFundos deverá ser sempre conferida e atualizada ao perceber que o PROCV dos CNPJs não retorna nenhum resultado ou erro, sinalizando que um novo fundo foi inserido no esquema

A planilha dClientes deverá ser sempre conferida e atualizada conforme o número de clientes aumente ou haja mudanças nos assessores

Ao ser aberta a planilha, deve-se logar no Add-In do Economática e esperar a atualização dos dados para que estes possam se refletir no Power BI

Nos dias de atualização do relatório, deve-se ter em mãos a versão mais recente do arquivo "visao_operacoes" do Safra Invest e copiá-la para 'Base Molde' e depois transpô-la para 'Base Safra', seguindo-se os critérios descritos

![image](https://github.com/viniciusfjacinto/AAI_Investimentos/assets/87664450/05074f9d-a782-4b28-a8b7-1dcd86c43364)

# 2 - Script em R para calcular o Retorno das Carteiras em D-1 / A Volatilidade destas tanto na Data de Aportes quando em D-1 / Possibilitando a análise de variação

Ainda, são feitas análises adicionais, como a Volatilidade Ewma das Carteiras, o cálculo do Retorno considerando fundos resgatados que geraram lucro ou prejuízo

O script foi comentado e dividido em sub-seções para facilitar seu entendimento. Além do mais, os resultados foram separados em outputs essenciais que são consumidos pelo Power BI (tabelas denominadas T1, T2, T3 e T4)

#### T1 = Volatilidade do Portfolio na Data de Aplicacaoo, incluindo Vol Individual, Peso e Wnon

#### T2 = Volatilidade do Portfolio em D-1, incluindo Vol Individual, Peso e Wnon

#### T3 = Variacao nas Volatilidades do Portfolio 

#### T4 = Rentabilidades em D-1 pelos dois métodos

Foi feita uma divisão manual do Script em partes menores que foram introduzidas diretamente em consultas no Power Query pelo Power BI. Estas são acionadas toda vez que o relatório se atualizar e, para que o uso do script dentro da ferramenta fosse possível, foi necessário remover alguns pacotes como Knitr, KableExtra, Writexlsx, GGPlot2, MailsendR, além de remover o código de tabelas, gráficos e salvamento de arquivos

Com isso, a função do R dentro do projeto consiste em apenas debugar algum problema futuro que venha a surgir e realizar o envio de e-mails sazonais

Sobre isso, e-mails sazonais serão enviados para os assessores indicando quais clientes necessitam de atenção por apresentarem variações na Volatilidade da Carteira acima ou abaixo dos parâmetros estimados. Tudo isso encontra-se programado em script auxiliar.

Exemplo de Report para um Cliente - Caso a volatilidade da carteira (simples ou EWMA) ultrapasse as bandas em roxo, indicando uma variação maior que 30%, então os ativos deverão ser rebalanceados.

![image](https://github.com/viniciusfjacinto/AAI_Investimentos/assets/87664450/d60fbab4-b5d6-435a-8601-8c4b28ceffbd)


Exemplo de Análise dos Retornos EWMA (Diário) para um Cliente

![image](https://github.com/viniciusfjacinto/AAI_Investimentos/assets/87664450/f1c85bb8-3629-4abb-841b-1f8a38fb16ac)


# 3 -Power BI

A última etapa consistiu no desenvolvimento dos visuais e medidas necessárias à visualização dos dados nos formatos necessários, proporcionando a agilidade e a organização requeridos para o trabalho no dia a dia

Todas as medidas foram dispostas dentro da tabela _Medidas

Foi feita uma consulta no CRM da Empresa utilizando uma API Get diretamente pelo Power Query para completar os dados de clientes com informações qualitativas

As consultas foram separads em pastass (R, Excel e Pipedrive), para detonar se são provenientes de um Script em R, da base em Excel ou da API do CRM, sendo que as menos importantes, ou aquelas que apenas servem para merge/append, foram ocultadas

Ainda, utilizamos o arquivo Lista de Fundos_MÊS para filtrar aqueles Fundos abertos e fechados, que também norteiam parte do trabalho cotidiano na empresa

![image](https://github.com/viniciusfjacinto/AAI_Investimentos/assets/87664450/e41cda1e-2788-4405-b103-9efa36504412)
![image](https://github.com/viniciusfjacinto/AAI_Investimentos/assets/87664450/17b9d352-5597-4693-a0c4-413a97f58718)

# 4 - Clustering

Como estudo, foi feito um clustering dos fundos que permitiu complementar a análise da matriz de correlação para a seleção de Fundos com base em suas variâncias conjuntas e, também, com base em seus agrupamentos categórios de retorno de 1 dia, retorno de 1 ano, volatilidade, valor da cota, sharpe e sortino. O agrupamento foi feito pelos métodos WARD e KMEANS.

Idealmente, cada carteira deve se especializar em uma categoria que seja aderente ao perfil do investidor, ou conter ao menos 1 fundo de cada grupo para diverisifcação.

![image](https://github.com/viniciusfjacinto/AAI_Investimentos/assets/87664450/91d61ec0-e93e-43cd-a5c6-31f62273bf6c)


