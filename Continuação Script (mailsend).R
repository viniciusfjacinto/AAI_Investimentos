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
               username = "viniciusfjacinto@gmail.com",
               password = "vj110796")


#Cria uma lista unica com cada assessor
Assessores <- list(unique(T3VolAlta$Assessor))


# Enzo ----------------------------------------------------------------------------------------

#Prepara Email Enzo
T3_Enzo <- filter(T3VolAlta, T3VolAlta$Assessor == Assessores[[1]][[2]])
path_enzo <- tempfile(fileext = ".xlsx")
write.xlsx(T3_Enzo, path_enzo)
T3_Email_Enzo <- "enzo.constantino@hotmail.com"

#Envia Email Enzo
email <- envelope() %>%
  attachment(path_enzo) %>% 
  from("viniciusfjacinto@gmail.com") %>%
  to(T3_Email_Enzo) %>%
  subject("Alertas de Volatilidade das Carteiras") %>% 
  text("Bom dia! 
Segue a relacao de clientes que apresentaram uma variacao da volatilidade maior que 30%. Para mais detalhes, acesse a Aba 'Portfolio Clientes' e clique em 'Visao Completa' no dashboard do link: 
https://app.powerbi.com/view?r=eyJrIjoiNGQ0MjBmY2YtNzZjNi00NWI0LTk4Y2YtNGNhYzNlYTEzMmE4IiwidCI6ImM3Mzk2ZTVlLTYzMDYtNGIwZi1hN2NmLWI1YzFhNDRkNDk0MSJ9&pageName=ReportSectionea01334fd03fe79a4dc5
       
*Obs: Todos os valores estao em %")

smtp(email, verbose = TRUE)


# Leo -----------------------------------------------------------------------------------------


#Prepara Email Leo
T3_Leo <- T3VolAlta
path_leo <- tempfile(fileext = ".xlsx")
write.xlsx(T3_Leo, path_leo)
T3_Email_Leo <- "leopac@hotmail.com"

#Envia Email Leo
email <- envelope() %>%
  attachment(path_leo) %>% 
  from("viniciusfjacinto@gmail.com") %>%
  to(T3_Email_Leo) %>%
  subject("Alertas de Volatilidade das Carteiras") %>% 
  text("Bom dia! 
Segue a relacao de clientes que apresentaram uma variacao da volatilidade maior que 30%. Para mais detalhes, acesse a Aba 'Portfolio Clientes' e clique em 'Visao Completa' no dashboard do link: 
https://app.powerbi.com/view?r=eyJrIjoiNGQ0MjBmY2YtNzZjNi00NWI0LTk4Y2YtNGNhYzNlYTEzMmE4IiwidCI6ImM3Mzk2ZTVlLTYzMDYtNGIwZi1hN2NmLWI1YzFhNDRkNDk0MSJ9&pageName=ReportSectionea01334fd03fe79a4dc5
       
*Obs: Todos os valores estao em %")
smtp(email, verbose = TRUE)
