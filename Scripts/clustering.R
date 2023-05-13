library(imputeTS)
library(janitor)
library(tidyverse) #pacote para manipulacao de dados
library(cluster) #algoritmo de cluster
library(dendextend) #compara dendogramas
library(factoextra) #algoritmo de cluster e visualizacao
library(fpc) #algoritmo de cluster e visualizacao
library(gridExtra) #para a funcao grid arrange
library(readxl)
library(PerformanceAnalytics)
library(plotly)
library(knitr)
library(kableExtra)

Fundos5orig <- read_excel("C:\\Users\\Cautela\\Desktop\\BI\\EntreRios\\Economática\\Fundos5.xlsx")
Fundos5orig <- Fundos5orig[-1,]




Fundos5 <- Fundos5orig[,-c(1,3:5)]
Fundos5[,2:10] <- lapply(Fundos5[,2:10],as.numeric)
Fundos5 <- Fundos5 %>%  na_replace(fill = 0)
Fundos5 <- column_to_rownames(Fundos5, var = "Nome")

Fundos5 <- Fundos5[,-c(4:6)]

####REMOVENDO OUTLIERS
Fundos5 <- Fundos5 %>% filter(`Valor da Cota mais recente` < 4000)

corr <- cor(Fundos5) 

#chart.Correlation(Fundos5)

fundospad <- scale(Fundos5)

fviz_nbclust(fundospad, FUN = hcut, method = "wss")

fviz_nbclust(fundospad, FUN = hcut, method = "silhouette")

#clusters <- 7

#####
dfundos <- dist(Fundos5, method = "euclidean")

#fundoshc1 <- hclust(dfundos, method = "single" )
#plot(fundoshc1, cex = 0.6, hang = -1)
#rect.hclust(fundoshc1, k = 7)

#dfundoscut <- cutree(fundoshc1,k = 7)
#table(dfundoscut)

####
fundoshc4 <- hclust(dfundos, method = "ward.D" )
#plot(fundoshc4, cex = 0.6, hang = -1)
#rect.hclust(fundoshc4, k = 6)

dfundoscut2 <- cutree(fundoshc4,k = 6)
table(dfundoscut2)

##############
fundos_cut_df <- dfundoscut2 %>%  as.data.frame()
fundos_fim <- cbind(Fundos5, fundos_cut_df)

mediagrupo_fundos <- fundos_fim[,1:7] %>% 
  group_by(fundos_fim$.) %>% 
  summarise(n = n(),
            retorno1d = mean(`Retorno do fechamento em 1 dia`),
            retorno1y = mean(`Retorno do fechamento em 1 ano`),
            cotalatest = mean(`Valor da Cota mais recente`),
            vol = mean(`Volatilidade base anual`),
            sortin = mean(Sortino),
            sharp = mean(Sharpe)
  )
  
mediagrupo_fundos %>%  
  kable %>%  
  kable_styling(
    bootstrap_options = "basic",
    font_size = "18",
    position = "center",
    stripe_color = "gray!6"
)
            
########
#K MEANS

fviz_nbclust(fundospad, kmeans, method = "wss")

fundosk2 <- kmeans(fundospad, centers = 2)
fundosk5 <- kmeans(fundospad, centers = 5)
fundosk7 <- kmeans(fundospad, centers = 7)

G1 <- fviz_cluster(fundosk2, geom = "point", data = fundospad) + ggtitle("k=2")
G2 <- fviz_cluster(fundosk5, geom = "point", data = fundospad) + ggtitle("k=5")
G3 <- fviz_cluster(fundosk7, geom = "point", data = fundospad) + ggtitle("k=7")

grid.arrange(G1,G2,G3, nrow = 2)

fundos_fim <- cbind(fundos_fim, fundosk5$cluster)

mediagrupo_fundoskmeans <- fundos_fim[,c(1:6,8)] %>% 
  group_by(fundos_fim$`fundosk5$cluster`) %>% 
  summarise(n = n(),
            retorno1d = mean(`Retorno do fechamento em 1 dia`),
            retorno1y = mean(`Retorno do fechamento em 1 ano`),
            cotalatest = mean(`Valor da Cota mais recente`),
            vol = mean(`Volatilidade base anual`),
            sortin = mean(Sortino),
            sharp = mean(Sharpe)
  )

mediagrupo_fundoskmeans %>%  
  kable %>%  
  kable_styling(
    bootstrap_options = "basic",
    font_size = "18",
    position = "center",
    stripe_color = "gray!6"
  )

fundos_fim <- rename(fundos_fim, "WARD" = .)
fundos_fim <- rename(fundos_fim, "KMEANS" = "fundosk5$cluster")

library(xlsx)
folderA = "c:/Users/Cautela/Desktop"
#xlsx::write.xlsx(fundos_fim, file = "c:/Users/Cautela/Desktop/clustfundos.xlsx")

mediagrupo_fundos <- rename(mediagrupo_fundos, "Agrup" = "fundos_fim$.")
mediagrupo_fundos$type <- "WARD" 

mediagrupo_fundoskmeans <- rename(mediagrupo_fundoskmeans, "Agrup" ="fundos_fim$`fundosk5$cluster`")
mediagrupo_fundoskmeans$type <- "KMEANS"

library(compareDF)
library(htmlTable)

ctable_medias <- compare_df(mediagrupo_fundos,mediagrupo_fundoskmeans)

#compareDF::create_output_table(ctable_medias)

ctable_medias$comparison_df %>%
  arrange(type) %>%
  mutate(chng_type = NULL, Agrup = NULL) %>% 
  kable (escape = FALSE) %>%
  kable_styling (bootstrap_options =  "basic")

ctable_medias$comparison_df %>%
  mutate(Agrup = NULL) %>% 
  mutate(chng_type = 
           cell_spec(chng_type, color = ifelse(chng_type == "+", "green", "red"))) %>% 
  kable (escape = FALSE) %>%
  kable_styling (bootstrap_options =  "basic") %>% 
  column_spec(2, bold = TRUE, color = "black", background = "lightgrey") 

options ( scipen = 999)
FundosOut <- Fundos5orig %>% lapply(as.numeric)
plot(Fundos5$`Valor da Cota mais recente`, xlim=c(0,680), ylim=c(0,50000),main = "Com outliers", pch = "x", col = "red")
abline(h = 2*mean(Fundos5$`Valor da Cota mais recente`, na.rm = T), col = "blue")

library(outliers)
library(ggplot2)

fundos_fim$WARD <- fundos_fim$WARD %>% as.factor()
fundos_fim$KMEANS <- fundos_fim$KMEANS %>% as.factor()


#2d PLOT
ggplotly(
ggplot(fundos_fim,
       aes(
       x = fundos_fim$`Valor da Cota mais recente`,
       y = fundos_fim$`Retorno do fechamento em 1 ano`,
       color = fundos_fim$KMEANS)) + 
  geom_point() +
  xlim(0,1500) +
  ylim(0,60) 
)

ggplotly(
  ggplot(fundos_fim,
         aes(
           x = fundos_fim$`Valor da Cota mais recente`,
           y = fundos_fim$`Retorno do fechamento em 1 ano`,
           color = fundos_fim$WARD)) + 
    geom_point() +
    xlim(0,1500) +
    ylim(0,40) 
)

#3D Plot

plot_kmeans <- plot_ly(fundos_fim, x = ~`Valor da Cota mais recente`, y = ~`Retorno do fechamento em 1 ano`, z = ~`Volatilidade base anual`, color = ~KMEANS) %>%
  add_markers(size = 1) %>%
  layout(scene = list(xaxis = list(title = "Valor da Cota mais recente"),
                      yaxis = list(title = "Retorno do fechamento em 1 ano"),
                      zaxis = list(title = "Volatilidade Base Anual")),
         title = "KMEANS Clustering")

plot_ward <- plot_ly(fundos_fim, x = ~`Valor da Cota mais recente`, y = ~`Retorno do fechamento em 1 ano`, z = ~`Volatilidade base anual`, color = ~WARD) %>%
  add_markers(size = 1) %>%
  layout(scene = list(xaxis = list(title = "Valor da Cota mais recente"),
                      yaxis = list(title = "Retorno do fechamento em 1 ano"),
                      zaxis = list(title = "Volatilidade Base Anual")),
         title = "WARD Clustering")

subplot(plot_kmeans,nrows = 1)


