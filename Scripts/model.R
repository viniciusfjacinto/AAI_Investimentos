# #Carregando os pacotes necessarios ----------------------------------------------------------


pacotes <- c("write_xlsx","writexl","dplyr","knitr", "kableExtra","imputeTS",
             "readxl","janitor","data.table","lubridate","bizdays","tidyquant",
             "xlsx","purrr","tidyverse")

if(sum(as.numeric(!pacotes %in% installed.packages())) != 0){
  instalador <- pacotes[!pacotes %in% installed.packages()]
  for(i in 1:length(instalador)) {
    install.packages(instalador, dependencies = T)
    break()}
  sapply(pacotes, require, character = T) 
} else {
  sapply(pacotes, require, character = T) 
}

#Customizacoes
options(scipen=999)
options(digits = 4)


# Introducao - Carregando a Base e determinando os subprodutos  ---------------------------------------------------------------------------


##Notas Importantes - Objetos terminados em In se referem a Carteira Inicial
## Objetos terminados em Fn se referem a carteira na data D-1 (Ontem)

##Objetos denominados T1, T2, T3... se referem aos produtos finais do Algoritmo
## T1 = Volatilidade do Portfolio na Data de Aplicacaoo, incluindo Vol Individual, Peso e Wnon
## T2 = Volatilidade do Portfolio em D-1, incluindo Vol Individual, Peso e Wnon
## T3 = Variacao nas Volatilidades do Portfolio 
## TabelasRent = Ganhos comparados na Data Aplicacao e em D-1
## Rentabilidade = Rentabilidade integral considerando ganhos e prejuizos com resgates totais
## RentabilidadeFAtivos = Rentabilidade que considera apenas os fundos ativos

#Pasta dos Arquivos
FolderA <- "c:/Users/Cautela/Desktop/EntreRios/Economática/"

#Carrega base Fundos
Fundos <- read_excel(paste0(FolderA,"Fundos5.xlsx"))
#Remove Ticker
Fundos <- Fundos[-1,]
#Volatilidade de character para numeric
Fundos$`Volatilidade base anual` <- as.numeric(Fundos$`Volatilidade base anual`)

#Carrega base de Retornos
FundosRt <- read_excel(paste0(FolderA,"Fundos5.xlsx"), 
                       sheet = "Retorno Fundos 1Ydaybyday")

###############################
FundosRt <- FundosRt[c(1:5,ncol(FundosRt):6)]
###############################

#Carrega base de Clientes com as infos na data de aplicacao
BaseSafra <- read_excel(paste0(FolderA,"Fundos5.xlsx"), 
                        sheet = "Base Safra") %>%  remove_empty("cols")
#Datas transformadas em formato Date
BaseSafra$Data <- BaseSafra$Data %>% as.Date()

#Inverte a ordem dos retornos
#BaseSafra<- BaseSafra #[c(1:10,ncol(BaseSafra):11)]

#Substitui NA's por 0
BaseSafra <- BaseSafra %>%  na_replace(fill = 0)


# Preparando os objetos -----------------------------------------------------------------------


#Lista com Clientes
Clientes <- list(unique(BaseSafra$Cliente))
#Transforma em Vetor para possibilitar o uso no For Loop
VClientes <- unlist(Clientes,use.names = TRUE)

#Isola as informacoees da planilha de Operacoes
OperacoesInfo <- BaseSafra[,1:7]

#Substitui NA's por 0
OperacoesInfo <- na_replace(OperacoesInfo, fill=0)

#Transforma CNPJ's de character para numeric, possibilitando o MERGE
OperacoesInfo$CNPJ <- as.double(OperacoesInfo$CNPJ)
Fundos$CNPJ <- as.double(Fundos$CNPJ)

#Merge para trazer a volatilidade atual
OperacoesInfoFn <- left_join(OperacoesInfo , (Fundos[,c('CNPJ','Volatilidade base anual')]),
                             by = 'CNPJ')

#Isola as informacoes da planilha de Operacoes com a Volatilidade na data de aplicacao
#Remove coluna vazia
OperacoesInfoIn <- BaseSafra[,1:9]

#Separa o retorno dos fundos na data de aplicacao para calculo da Matriz de Correlacao
OperacoesInfoInRt <- BaseSafra %>%  select(-`Cota D+1`)
#Transforma CNPJ em numerico para possibilitar a aplicacao da funcao COR
OperacoesInfoInRt$CNPJ <- as.numeric(OperacoesInfoInRt$CNPJ)


# Calculando as Carteiras Iniciais e Volatilidade ------------------------------------------------------------


##FOR LOOP - Declarando variaveis
TabelasIn = NULL; TabelasInReturn = NULL; TabelasInCorr = NULL; 
MatrixIn = NULL; MatrixIn_1 = NULL; MatrixWNon = NULL; 
MatrixWNonT = NULL; VolPortIn = NULL

##FOR LOOP - Carteira Inicial
for (i in 1:length(VClientes)) {
  
  #Cria uma tabela para cada cliente, separadamente
  #Filtra a primeira data + 5 dias uteis para estabelecer o aporte inicial
  #Filtra so as aplicacoes
  TabelasIn[[i]] <- OperacoesInfoInRt %>%  
    filter(Cliente == VClientes[[i]]) %>% 
    #list() %>% 
    as.data.frame %>% 
    group_by(Cliente) %>% 
    filter(Data %in% min(Data):add.bizdays(min(Data),5,"Brazil/ANBIMA")) %>% 
    filter(Natureza == "Aplicação")
  
  #Calcula Peso e Volatilidade
  TabelasIn[[i]]$Peso <- TabelasIn[[i]]$Soma / sum(TabelasIn[[i]]$Soma)
  TabelasIn[[i]]$WVol <- TabelasIn[[i]]$Peso * (TabelasIn[[i]]$Volatilidade / 100)
  TabelasIn[[i]] <- TabelasIn[[i]] %>%  na_replace(fill=0) %>%  relocate(Peso,WVol, .after = Cota)
  
  #Cria novo objeto apenas com colunas CNPJ e Retornos
  TabelasInReturn[[i]] <- TabelasIn[[i]] %>%  select(-one_of(
    'WVol','Peso','Data','Mercadoria/Fundo',
    'Natureza','Volatilidade','Cota','Soma','Quantidade')
  )
  TabelasInReturn[[i]] <- TabelasInReturn[[i]][,-1] %>%  t()
  colnames(TabelasInReturn[[i]]) <- TabelasInReturn[[i]][1,]
  TabelasInReturn[[i]] <- TabelasInReturn[[i]][-1,]
  
  #Calcula as correlacoes para os fundos de cada cliente
  TabelasInCorr[[i]] <- as.matrix(TabelasInReturn[[i]])
  TabelasInCorr[[i]] <- cor(TabelasInCorr[[i]])
  
  #Junta as matrizes de correlacao com as tabelas do cliente
  #MatrixIn[[i]] <- cbind(x1 = TabelasIn[[i]][,-12:-377],as.data.frame(t(TabelasInCorr[[i]])))
  MatrixIn[[i]] <- cbind(x1 = TabelasIn[[i]][,-12:-272],as.data.frame(t(TabelasInCorr[[i]])))
  
  #Remove as demais informacoes e transforma em matriz
  MatrixIn_1[[i]] <- MatrixIn[[i]][,-1:-11] %>%  
    na_replace(fill=0) %>% 
    as.matrix()
  #Substitui NA's por 0
  MatrixWNon[[i]] <- MatrixIn[[i]]$WVol %>% 
    na_replace(fill=0) %>% 
    as.matrix()
  #Cria uma nova Matriz transposta
  MatrixWNonT[[i]] <- t(MatrixWNon[[i]]) %>%  
    as.matrix()
  
  #Calculo da Volatilidade do Portfolio
  VolPortIn[[i]] <- sqrt((MatrixWNonT[[i]] %*% MatrixIn_1[[i]]) %*% MatrixWNon[[i]])
  
  #Une os dados novamente com a tabela dos clientes
  VolPortIn[[i]] <- cbind(x1 = MatrixIn[[i]][,1:11] , as.data.frame(VolPortIn[[i]]))
}

#Primeiro Resultado - T1
#Junta todas as tabelas novamente
T1 <- rbindlist(VolPortIn)

#Renomeia V1 para Volatilidade Portfolio
T1 <- rename(T1,"Volatilidade Portfolio" = V1)



# Calculo do Retorno e Volatilidade das Carteiras em D-1 -----------------------------------------------------


##OperacoesInfoFn - para calculo da Volatilidade

#OperacoesCotaFn - para calculo da Rentabilidade
OperacoesCotaFn <- BaseSafra[,1:10]
OperacoesCotaFn <- OperacoesCotaFn %>%  select(-'Volatilidade')
OperacoesCotaFn$CNPJ <- OperacoesCotaFn$CNPJ %>% as.numeric
OperacoesCotaFn <- left_join(OperacoesCotaFn,
                             Fundos[,c('CNPJ','Valor da Cota mais recente')],
                             by = 'CNPJ')
OperacoesCotaFn <- rename(OperacoesCotaFn, "CotaInicial" = Cota,
                          "CotaInicialD1" = 'Cota D+1',
                          "CotaLatest" = "Valor da Cota mais recente")
OperacoesCotaFn$CotaLatest <- OperacoesCotaFn$CotaLatest %>%  as.numeric()

#################################
VClientesT1 <- unique(T1$Cliente)
#################################

#Declara variaveis para o Loop
TabelasFn = NULL; TabelasFnGroup = NULL;
TabelasFnGroupRt = NULL;TabelasFnCorr = NULL;
MatrixFn = NULL; MatrixFn_1 = NULL; MatrixFnWNon = NULL;
MatrixFnWNonT = NULL; VolPortFn = NULL;
TabelasFnRentAtiva = NULL;

Ewma_Vol = NULL; Ewma_VolDiff = NULL; Ewma_Corr = NULL;
Ewma_Matriz = NULL; Ewma_Matriz_1 = NULL; Ewma_MatrizWnon = NULL;
Ewma_MatrizWnonT = NULL; Ewma_VolPort = NULL;
dx = NULL

lambda = 0.94

#Carteira Hoje
for (i in 1:length(VClientesT1)) {
  TabelasFn[[i]] <- OperacoesCotaFn %>%  
    filter(Cliente == VClientesT1[[i]]) %>% 
    list() %>% 
    as.data.frame %>% 
    group_by(Cliente)
  
  TabelasFn[[i]]$Data <- as.Date(TabelasFn[[i]]$Data)
  
  #Cria uma coluna com as Cotas Iniciais, obedecendo algumas regras
  #Regra1: Cota Inicial Manual = Valor/Qtd, se Aplicacao
  #Regra2: Considera as diferencas entre Cota em D e Cota em D+1, e substitui
  #aquela que tiver a menor diferenca e ainda for <1 em comparacao com
  #a cota calculada manualmente, se nao mantem o valor de Regra1
  #Regra3: CotaInicialManual = CotaInicial se Resgate e Qtd <> 0
  #Regra4: |Valor|/Soma(Qtd) se Resgate e Qtd = 0
  #Regra5: Se |Cota Inicial Manual| - |Cota Inicial| > 20,
  #entao Cota Inicial Manual = Cota EconomÃ¡tica em D
  #Regra para corrigir um erro especifico em TabelasFn[[5]]
  
  #1
  TabelasFn[[i]]$CotaInicialManual <-
    TabelasFn[[i]]$Soma / TabelasFn[[i]]$Quantidade
  
  #2
  TabelasFn[[i]]$DifD <-
    abs(TabelasFn[[i]]$CotaInicial - TabelasFn[[i]]$CotaInicialManual)
  
  TabelasFn[[i]]$DifD1 <-
    abs(TabelasFn[[i]]$CotaInicialD1 - TabelasFn[[i]]$CotaInicialManual)
  
  TabelasFn[[i]]$CotaInicialManual <-
    ifelse(
      TabelasFn[[i]]$DifD < TabelasFn[[i]]$DifD1 & TabelasFn[[i]]$DifD < 1, 
      TabelasFn[[i]]$CotaInicial,
      ifelse(
        TabelasFn[[i]]$DifD1 < TabelasFn[[i]]$DifD & TabelasFn[[i]]$DifD1 < 1,
        TabelasFn[[i]]$CotaInicialD1,
        TabelasFn[[i]]$CotaInicialManual
      )
    )
  
  #4
  TabelasFn[[i]] <- TabelasFn[[i]] %>% group_by(Mercadoria.Fundo) %>%
    mutate(CotaInicialManual=case_when(Quantidade==0 ~ 
                                         abs(Soma)/sum(Quantidade, na.rm=TRUE)))
  
  TabelasFn[[i]]$CotaInicialManual <- ifelse(TabelasFn[[i]]$Quantidade == 0,
                                             TabelasFn[[i]]$CotaInicialManual, 
                                             TabelasFn[[i]]$Soma / TabelasFn[[i]]$Quantidade)
  #3
  TabelasFn[[i]]$CotaInicialManual <- 
    ifelse(TabelasFn[[i]]$Natureza %in% c('Resgate') & 
             TabelasFn[[i]]$Quantidade> 0,
           TabelasFn[[i]]$CotaInicial, TabelasFn[[i]]$CotaInicialManual)
  
  #5
  TabelasFn[[i]]$CotaInicialManual <-
    ifelse(
      (TabelasFn[[i]]$CotaInicial - TabelasFn[[i]]$CotaInicialManual) > abs(20),
      TabelasFn[[i]]$CotaInicial, TabelasFn[[i]]$CotaInicialManual    
    )
  
  #Remove as colunas ja utilizadas no processo de calculo
  TabelasFn[[i]] <- TabelasFn[[i]] %>%  select(-'DifD',-'DifD1',-'CotaInicial',-'CotaInicialD1')
  
  
  #Calcula a variacao entre as Cotas
  TabelasFn[[i]]$VarCota <- 
    (TabelasFn[[i]]$CotaLatest - TabelasFn[[i]]$CotaInicialManual) / TabelasFn[[i]]$CotaInicialManual
  
  #Calcula os rendimentos para cada Fundo
  TabelasFn[[i]]$Rendimentos <- TabelasFn[[i]]$VarCota * TabelasFn[[i]]$Soma
  
  #Calcula o Valor Final de cada fundo, ainda com duplicidades
  TabelasFn[[i]]$SomaFn <- TabelasFn[[i]]$Rendimentos + TabelasFn[[i]]$Soma
  
  #Calcula a Rentabilidade (REAL OU NOMINAL?)
  #TabelasFn[[i]]$Rentabilidade <- 
  #(sum(TabelasFn[[i]]$SomaFn)-sum(TabelasFn[[i]]$Soma))/sum(TabelasFn[[i]]$Soma)
  
  TabelasFn[[i]]<- TabelasFn[[i]] %>%  
    group_by(Cliente) %>% 
    mutate(Rentabilidade = 
             sum(SomaFn*(Quantidade>0) + 
                   Rendimentos * (Quantidade==0)) /
             sum(Soma * (Quantidade > 0)) - 1) %>%   
    as.data.frame()
  
  
  
  #Agrega os Fundos de mesmo nome pelo Valor Final
  TabelasFnGroup[[i]] <- 
    aggregate(TabelasFn[[i]]$SomaFn, 
              by = list(Fundo = TabelasFn[[i]]$Mercadoria.Fundo), 
              FUN = sum)
  
  #TabelasFnGroupTest[[i]] <- aggregate(TabelasFn[[i]]$Soma,
  #by = list(Fundo = TabelasFn[[i]]$Mercadoria.Fundo), FUN = sum)
  
  #Junta as duas tabelas
  TabelasFnGroup[[i]] <- merge(TabelasFn[[i]],
                               TabelasFnGroup[[i]], 
                               by.x = "Mercadoria.Fundo", 
                               by.y = "Fundo")
  
  
  #Remove duplicatas
  TabelasFnGroup[[i]] <- TabelasFnGroup[[i]][!duplicated(TabelasFnGroup[[i]][,c("x")]),]
  
  
  
  #Renomeia Coluna de Fundos e remove colunas indesejadas
  TabelasFnGroup[[i]] <- rename(TabelasFnGroup[[i]],"ValorFinal"=x)
  TabelasFnGroup[[i]] <- TabelasFnGroup[[i]][,c(-4:-6,-8:-13)]
  
  #Remove residuo de fundos que foram retirados da Carteira
  TabelasFnGroup[[i]] <- TabelasFnGroup[[i]][TabelasFnGroup[[i]]$ValorFinal > 20,]
  
  #Preparo para calculo da Volatilidade do Portfolio
  #Traz a volatilidade dos fundos
  TabelasFnGroup[[i]] <- left_join(TabelasFnGroup[[i]],
                                   Fundos[,c('CNPJ','Volatilidade base anual')],
                                   by = 'CNPJ')
  
  #Calcula Peso e WVol
  TabelasFnGroup[[i]]$Peso <- 
    TabelasFnGroup[[i]]$ValorFinal / sum(TabelasFnGroup[[i]]$ValorFinal)
  
  TabelasFnGroup[[i]]$WVol <- 
    TabelasFnGroup[[i]]$Peso * (TabelasFnGroup[[i]]$`Volatilidade base anual` / 100)
  
  #left_join entre os valores finais e os retornos diarios
  TabelasFnGroup[[i]] <- left_join(TabelasFnGroup[[i]],
                                   select(FundosRt,'CNPJ',c(6:268)), by = 'CNPJ')
  
  #Isola Retorno e CNPJ
  TabelasFnGroupRt[[i]] <- TabelasFnGroup[[i]] %>%  select(-one_of(
    'WVol','Peso','Data','Mercadoria.Fundo',
    'Volatilidade base anual','ValorFinal'))
  
  #Transpoe e remove clientes
  TabelasFnGroupRt[[i]] <- TabelasFnGroupRt[[i]][,-1] %>% 
    na_replace(fill=0) %>%
    t()
  colnames(TabelasFnGroupRt[[i]]) <- TabelasFnGroupRt[[i]][1,]
  TabelasFnGroupRt[[i]] <- TabelasFnGroupRt[[i]][-1,]
  
  #Calcula a Matriz de correlacao dos valores finais
  TabelasFnCorr[[i]] <- cor(as.matrix(TabelasFnGroupRt[[i]]))
  
  #Junta as matrizes de correlacao com as tabelas do cliente
  MatrixFn[[i]] <- cbind(TabelasFnGroup[[i]][,-9:-271],as.data.frame(t(TabelasFnCorr[[i]])))
  
  #Remove as demais informacoes e transforma em matriz
  MatrixFn_1[[i]] <- MatrixFn[[i]][,-1:-8] %>%  
    na_replace(fill=0) %>% 
    as.matrix()
  
  #Substitui NA's por 0
  MatrixFnWNon[[i]] <- MatrixFn[[i]]$WVol %>% 
    na_replace(fill=0) %>% 
    as.matrix()
  
  #Cria uma nova Matriz transposta
  MatrixFnWNonT[[i]] <- 
    t(MatrixFnWNon[[i]]) %>%  
    as.matrix()

  #Calculo da Volatilidade do Portfolio
  VolPortFn[[i]] <- 
    sqrt((MatrixFnWNonT[[i]] %*% MatrixFn_1[[i]]) %*% MatrixFnWNon[[i]]) %>%
    as.data.frame()
 
  #Une os dados novamente com a tabela dos clientes
  VolPortFn[[i]] <- 
    cbind(MatrixFn[[i]][,1:8] , VolPortFn[[i]])


  ###Calcula Volatilidade EWMA
  
  Ewma_Vol[[i]] <- TabelasFnGroupRt[[i]] %>% as.data.frame()
  
  Ewma_Vol[[i]] <- as.matrix(Ewma_Vol[[i]][,1:ncol(Ewma_Vol[[i]])])^2
  
  Ewma_VolDiff [[i]] <- sapply(data.frame(Ewma_Vol[[i]]),diff)
  
  dx <- outer(1:nrow(Ewma_VolDiff[[i]]), 1:nrow(Ewma_VolDiff[[i]]),
              FUN = function(x,y)
                ifelse (x >= y, lambda^(x-y+1), 0 )
  )
  
  Ewma_Vol[[i]] <- Ewma_Vol[[i]] - rbind(0, dx %*% Ewma_VolDiff[[i]])
  
  Ewma_Corr[[i]] <- cor(as.matrix(Ewma_Vol[[i]]))
  
  Ewma_Matriz[[i]] <- cbind(TabelasFnGroup[[i]][,-9:-271],as.data.frame(t(Ewma_Corr[[i]])))
  
  Ewma_Matriz_1[[i]] <- Ewma_Matriz[[i]][,-1:-8] %>%  
    na_replace(fill=0) %>% 
    as.matrix()
  
  Ewma_MatrizWnon[[i]] <- Ewma_Matriz[[i]]$WVol %>% 
    na_replace(fill=0) %>% 
    as.matrix()
  
  Ewma_MatrizWnonT[[i]] <- t(Ewma_MatrizWnon[[i]]) %>% 
    as.matrix()
  
  Ewma_VolPort[[i]] <- 
    sqrt(Ewma_MatrizWnonT[[i]] %*% Ewma_Matriz_1[[i]] %*% Ewma_MatrizWnon[[i]])
  
  Ewma_VolPort[[i]] <-
    cbind(Ewma_Matriz[[i]][,1:8], as.data.frame(Ewma_VolPort[[i]]))
}  

#Segundo Resultado - T2
#Junta todas as tabelas novamente
T2 <- 
  rbindlist(VolPortFn)

#Renomeia V1 para Volatilidade Portfolio
T2 <- 
  rename(T2,"Volatilidade Portfolio D" = V1)

#write_xlsx(T2, path = "CarteiraClienteHoje.xlsx")

#Segundo Resultado Parcial - T2 Ewma (declara e renomeia)
T2Ewma <- rbindlist(Ewma_VolPort) %>%  select("Cliente","V1")
T2Ewma <- 
  rename(T2Ewma,"Volatilidade Portfolio Ewma em D" = V1)


# SUBPRODUTOS PARA POWERBI --------------------------------------------------------------------


#1 -  Informacoes completas com Rentabilidade
TabelasRent <- rbindlist(TabelasFn)

#2 - Rentabilidade com fundos inativos
Rentabilidade <- rbindlist(TabelasFn) %>% 
  select(c('Cliente','Rentabilidade')) %>% 
  unique()

#3 - Rentabilidade com fundos ativos
RentabilidadeFAtivos <- rbindlist(TabelasFn)

RentabilidadeFAtivos <- RentabilidadeFAtivos %>% 
  group_by(Cliente, Mercadoria.Fundo) %>% 
  arrange(Data) %>% 
  mutate(zero_point = case_when(Quantidade == 0 ~ 1)) %>%
  fill(zero_point, .direction = "up") %>% 
  filter(is.na(zero_point))

RentabilidadeFAtivos <- RentabilidadeFAtivos %>%
  select(c(-"zero_point"))

RentabilidadeFAtivos2 <- RentabilidadeFAtivos %>% 
  select(-"Data") %>% 
  select ("Cliente","Mercadoria.Fundo", "Soma", "SomaFn") 

RentabilidadeFAtivos3 <- RentabilidadeFAtivos2 %>% 
  group_by(Cliente,Mercadoria.Fundo) %>% 
  summarize(SomaFn = sum(SomaFn)) 

RentabilidadeFAtivos4 <- RentabilidadeFAtivos2 %>% 
  group_by(Cliente,Mercadoria.Fundo) %>% 
  summarize(Soma = sum(Soma)) 

RentabilidadeFAtivos2 <- cbind(RentabilidadeFAtivos3,RentabilidadeFAtivos4, by = "Cliente")
RentabilidadeFAtivos2 <- RentabilidadeFAtivos2[,c(-4,-5,-7)]
RentabilidadeFAtivos2 <- rename(RentabilidadeFAtivos2, "Cliente" = "Cliente...1")

RentabilidadeFAtivos2 <- RentabilidadeFAtivos2 %>%  group_by(Cliente) %>% 
  mutate(Rent = ((sum(SomaFn)/sum(Soma))-1))

#write.csv2(RentabilidadeFAtivos2, file = paste0(FolderA,"extrairR.csv"))

# Correlacao entre os Fundos ----------------------------------------------------------------------------------


##Calcular a matriz de correl de todos os fundos
FundosRtCorr <- 
  FundosRt[,!names(FundosRt) %in% c('Ticker','Nome','Empresa gestora','Classificação Anbima')]
FundosRtCorr <- t(FundosRtCorr)
colnames(FundosRtCorr) <- FundosRtCorr[1,]
FundosRtCorr <- FundosRtCorr[-1,]
FundosRtCorr <- na_replace(FundosRtCorr,fill=0)
FundosRtCorr <- cor(FundosRtCorr)
FundosRtCorr <- as.data.frame(FundosRtCorr)
FundosRtCorr = FundosRtCorr
FundosRtCorr <- cbind('CNPJ'= rownames(FundosRtCorr),FundosRtCorr)
rownames(FundosRtCorr) <- 1:nrow(FundosRtCorr)
FundosRtCorr$CNPJ <- as.numeric(FundosRtCorr$CNPJ)
FundosRtCorr <- (left_join(FundosRtCorr,FundosRt[,names(FundosRt) %in% c('Nome','CNPJ')], by = 'CNPJ'))
FundosRtCorr <- FundosRtCorr %>%  select('Nome', everything())

#write.csv(FundosRtCorr, file = paste0(FolderA,"extrairR.csv"))


# Analises e setup para os Alertas de Vol ------------------------------------------------------------------------------


#Analise Final Diaria para envio de e-mails, volatilidades > abs(30%)
#Cria a tabela com as Volatilidades por cliente (em D e na Data Aplicacao)
T3 <- left_join(unique(T1[,c("Cliente","Volatilidade Portfolio")]),unique(T2[,c("Cliente","Volatilidade Portfolio D")]), by = "Cliente")
T3$VarVol <- (T3$`Volatilidade Portfolio D` / T3$`Volatilidade Portfolio`) - 1

#Filtra aqueles clientes com VarVol > 30% ou < -30%
T3VolAlta <- filter(T3, abs(T3$VarVol)*100 > 30)

#Multiplica por 100 para deixar em porcentagem
T3VolAlta[,2:4] <- T3VolAlta[,2:4]*100

#Cria a tabela T3 na versao da Vol Ewma, e realiza os mesmos processos anteriores
T3Ewma <- left_join(unique(T1[,c("Cliente","Volatilidade Portfolio")]),
                    unique(T2Ewma[,c("Cliente","Volatilidade Portfolio Ewma em D")]), 
                    by = "Cliente")

T3Ewma$VarVol <- ((T3Ewma$`Volatilidade Portfolio Ewma em D`/T3$`Volatilidade Portfolio`)-1)
T3EwmaVolAlta <- filter(T3Ewma, abs(T3Ewma$VarVol)*100 > 30)
T3EwmaVolAlta[,2:4] <- T3EwmaVolAlta[,2:4]*100


#Visualiza as volatilidades comuns para os criterios filtrados (VarVol Comum)
T3VolAlta %>%
  kable() %>%
  kable_styling(bootstrap_options = "basic",
                full_width = F, 
                font_size = 20)

#Visualiza as volatilidades comuns para os criterios filtrados (VarVol Ewma)
T3EwmaVolAlta %>%
  kable() %>%
  kable_styling(bootstrap_options = "basic",
                full_width = F, 
                font_size = 20)

#Retorna a diferenca entre os clientes de atencao nos dois casos
count(T3EwmaVolAlta) - count(T3VolAlta) 

count(T3VolAlta)
count(T3EwmaVolAlta)

#Une as duas tabelas para analisar diferencas e quais clientes encontram-se
#em uma mas nao em outra

#Une as duas tabelas criadas anteriormente
VolAltaGeral <- left_join(T3EwmaVolAlta,T3, by = "Cliente")

#Multiplica os valores por 100 para termos o percentual
VolAltaGeral[,5:7] <- VolAltaGeral[,5:7]*100

#Remove e renomeia as colunas para melhorar a visualizacao da informacao
VolAltaGeral <- VolAltaGeral %>%  select(-`Volatilidade Portfolio.y`)
VolAltaGeral <- rename(VolAltaGeral, VarVolEwma = "VarVol.x")
VolAltaGeral <- rename(VolAltaGeral, VarVolComum = "VarVol.y")
VolAltaGeral <- rename(VolAltaGeral, "Volatilidade Portfolio" = "Volatilidade Portfolio.x")

#Reorganiza a ordem das colunas
VolAltaGeral <- VolAltaGeral[,c(1,2,5,3,6,4)]

#Calcula a diferenca entre as Vols
VolAltaGeral$DiferencaVol <- 
  (VolAltaGeral$VarVolComum - VolAltaGeral$VarVolEwma)

#Exibe em uma tabela
VolAltaGeral %>%
  kable() %>%
  kable_styling(bootstrap_options = "basic",
                full_width = T, 
                font_size = 14)


write.xlsx(VolAltaGeral, file = paste0(FolderA,"VolComumxVolEwma.xlsx"))
