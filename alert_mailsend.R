# Carregamento dos Pacotes para Envio dos Emails ----------------------------------------------

install.packages('googlesheets4')
library(googlesheets4)
library(emayili)
library(magrittr)


# Carregamento da Base de Clientes ------------------------------------------------------------

dClientes <- read_sheet('https://docs.google.com/spreadsheets/d/1wQ5sSikEoirVsTyGAavVXeg0xwRhiNmyXtxlkhs4wdw/edit?usp=sharing')
dClientes$Cliente <- as.character(dClientes$Cliente)


# Modelagem dos Dados para possibilitar envio -------------------------------------------------


#Junta o alerta de Vol com os Assessores
T3VolAlta <- left_join(T3VolAlta, dClientes[,c("Cliente","Assessor")], by = "Cliente")

T3VolAlta %>%
  kable() %>%
  kable_styling(bootstrap_options = "basic",
                full_width = F, 
                font_size = 20)

#Nomeia o Servidor SMTP
smtp <- server(host = "smtp.gmail.com",
               port = 465,
               username = "aai_alert@gmail.com",
               password = "*********")


#Cria uma lista unica com cada assessor
Assessores <- list(unique(T3VolAlta$Assessor))


# Assessor 1 ----------------------------------------------------------------------------------------

#Prepara Email Assessor 1
T3_Enzo <- filter(T3VolAlta, T3VolAlta$Assessor == Assessores[[1]][[2]])
path_enzo <- tempfile(fileext = ".xlsx")
write.xlsx(T3_Assessor1, path_assessor1)
T3_Email_Enzo <- "assessor1@hotmail.com"

#Envia Email Assessor 1
email <- envelope() %>%
  attachment(path_enzo) %>% 
  from("alert_aai@gmail.com") %>%
  to(T3_Email_Assessor1) %>%
  subject("Alertas de Volatilidade das Carteiras") %>% 
  text("Bom dia! 
Segue a relacao de clientes que apresentaram uma variacao da volatilidade maior que 30%. Para mais detalhes, acesse a Aba 'Portfolio Clientes' e clique em 'Visao Completa' no dashboard do link: 
https://app.powerbi.com/view?r=eyJrIjoiNGQ0MjBmY2YtNzZjNi00NWI0LTk4Y2YtNGNhYzNlYTEzMmE4IiwidCI6ImM3Mzk2ZTVlLTYzMDYtNGIwZi1hN2NmLWI1YzFhNDRkNDk0MSJ9&pageName=ReportSectionea01334fd03fe79a4dc5
       
*Obs: Todos os valores estao em %")

smtp(email, verbose = TRUE)


# Gerente 1 -----------------------------------------------------------------------------------------


#Prepara Email Gerente 1
T3_Gerente1 <- T3VolAlta
path_gerente1 <- tempfile(fileext = ".xlsx")
write.xlsx(T3_Gerente1, path_gerente1)
T3_Email_Gerente1 <- "gerente1@hotmail.com"

#Envia Email Gerente 1
email <- envelope() %>%
  attachment(path_gerente1) %>% 
  from("alert_aai@gmail.com") %>%
  to(T3_Email_Gerente1) %>%
  subject("Alertas de Volatilidade das Carteiras") %>% 
  text("Bom dia! 
Segue a relacao de clientes que apresentaram uma variacao da volatilidade maior que 30%. Para mais detalhes, acesse a Aba 'Portfolio Clientes' e clique em 'Visao Completa' no dashboard do link: 
https://app.powerbi.com/view?r=eyJrIjoiNGQ0MjBmY2YtNzZjNi00NWI0LTk4Y2YtNGNhYzNlYTEzMmE4IiwidCI6ImM3Mzk2ZTVlLTYzMDYtNGIwZi1hN2NmLWI1YzFhNDRkNDk0MSJ9&pageName=ReportSectionea01334fd03fe79a4dc5
       
*Obs: Todos os valores estao em %")
smtp(email, verbose = TRUE)
