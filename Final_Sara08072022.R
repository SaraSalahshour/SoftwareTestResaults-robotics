install.packages("plotly")
install.packages("ncol")
install_github("ProcessMiner/nlcor")
install.packages("writexl")
install.packages("remotes")
remotes::install_github("Zelazny7/isofor")
remotes::install_github("https://github.com/talegari/solitude.git")
install.packages("ranger")
install.packages("rsample")
install.packages("naniar")
install.packages("ggfortify")
install.packages(" cranvas")
install.packages("qtbase")
devtools::install_github("tsieger/idendr0")
install.packages('dendextend')
install.packages("ecodist")
install.packages("NbClust")
install.packages("circlize")
install.packages("parameters")
install.packages('mclust')
install.packages('see')
install.packages('fpc')
install.packages("ClustOfVar", dependencies = TRUE) # dont need it
install.packages("pvclust")
install.packages("MVA")
install.packages("lsa")
install.packages("SnowballC")
install.packages("CCA")
install.packages("CCP")
install.packages("energy")
install.packages(" FactoMineR")
install.packages("candisc")
install.packages("clusterSim")
install.packages("ComputeSilhouetteScores")
install.packages("foregin")
install.packages("rstudioapi")
install.packages("shiny")
install.packages("glmnet")


library(NbClust)
library(tidyverse)
library(devtools)
library(ggpubr)
library(dplyr)
library(tidyr)
library(ggplot2)
library(gridExtra)
library(factoextra)
library(mice)
library(Hmisc)
library(corrplot)
library(devtools) 
library(plotly)
library(nlcor)
library(mice)
library(writexl)
library(isofor)
library(solitude)
library(ranger)
library("rsample")
library("solitude")
library("tidyverse")
library("mlbench")
library(caret)
library(naniar)
library(ggfortify)
library(MASS)
library(idendr0) #interactive dendrogram
library(dendextend)
library(ecodist)
library(circlize)
library(parameters)
library(mclust)
library(see)
library(cluster)
library(fpc)
library(ClustOfVar)
library(pvclust)
library(psych)
library(ellipse)
library(MVA)
library(lsa)
library(stylo)
library(CCA)
library(CCP)
library(energy)
library( FactoMineR)
library(candisc)
library(clusterSim)
library(ComputeSilhouetteScores)
library(foreign)
library(rstudioapi)
library(shiny)
library(glmnet)

###
###############
#CALLING DATA #
###############
# request the path to an existing .csv file on disk
path <- rstudioapi::selectFile(caption = "Select XLSX File",
                               filter = " (*.xlsx)",
                               existing = TRUE)

rstudioapi::showDialog(title = "Select sheet",
                       message = "<b>The default selection is Sheet=4. In line 108, you can type the required sheet number."
)
# now, you could read the data using e.g. 'readr::read_csv()'
Data <- readxl::read_xlsx(path, sheet = 6, skip = 4)
Data <- as.data.frame(Data)
row.names(Data) <- Data[,1]
Data <- Data[,-c(1,2)]
Data <- select_if(Data, is.numeric)             # Subset numeric columns with dplyr

###############
#CLEANING DATA#
###############
# read data from excel file under the name "ISO9283_v100_RW...EROP326Orange"
#ISO9283.N436 <- readxl::read_xlsx("C:/Users/RISE/Desktop/Sara/Internship/ABB/ISO9283-N436.xlsx" )
######### Cleaning data
## Remove columns with more than 75% missing data
iso.N436 <- Data[, which(colMeans(!is.na(Data)) > 0.75)]

## Remove columns with zero variances
df.iso <- data.frame(iso.N436[, sapply(iso.N436, function(x) length(unique(x)) > 2)])

#data.iso <- df.iso[,-c(1)] #remove Build column

# Drop Duplicated Columns:
df_dup <- df.iso[!duplicated(as.list(df.iso))] #dim df_dup 436*1939

#### Imputing missing data with mean
for(i in 1:ncol(df_dup)){
  df_dup[is.na(df_dup[,i]), i] <- mean(df_dup[,i], na.rm = TRUE)
}


### Delete column with correlation more than 0.97
corr.comp <- cor(df_dup)

# Modify correlation matrix
cor_matrix <- corr.comp                
cor_matrix[upper.tri(cor_matrix)] <- 0
diag(cor_matrix) <- 0

# Remove highly correlated variables
data_new <- df_dup[ , !apply(cor_matrix,    
                             2,
                             function(x) any(x >= 0.97))]

#row.names(data_new) <- df.iso[,1]


####################################
### Descriptive table and PLOT #####
####################################

# this function create table with Min, Mean, Median, Max and standard deviation
# it is also create plot for each variables

descriptive.table <- function(data,x) {
  myplot <- list()
  
  df <- as.data.frame(data %>% dplyr:: select(starts_with(x)))
  
  
  max_dat <- sapply(df, max, na.rm = TRUE) 
  min_dat <- sapply(df, min, na.rm = TRUE) 
  mean_dat <- sapply(df, mean, na.rm = TRUE) 
  sd_dat <- sapply(df, sd, na.rm = TRUE) 
  median_dat <- sapply(df, median, na.rm = TRUE) 
  
  df_descrip <- data.frame("Min"=min_dat,"Mean"=mean_dat,
                           "Median"= median_dat,"Max"=max_dat,"SD"=sd_dat)
  df_descrip
  
  list <- c("Mean","Median", "SD")
  cl <- rainbow(length(list))
  names(cl) <- list
  
  for(i in 1 : ncol(df)) {
    myplot[[i]] <- ggplot(data=data, aes_(x= c(1:nrow(data)), y=df[,i], na.rm = TRUE)) + geom_point(na.rm = TRUE) + 
      geom_hline(aes_(yintercept=mean(df[,i], na.rm = TRUE)), linetype="dashed", color = "red", size  = 1) +
      geom_hline(aes_(yintercept=median(df[,i], na.rm = TRUE)), linetype="dashed", color = "blue", size  = 1) +
      geom_hline(aes_(yintercept=quantile(df[,i], 0.90, na.rm = TRUE)), linetype="dashed", color = "green", size  = 1)+
      ggtitle(label = as.character(paste(colnames(df)[i])))+
      theme(axis.text.x=element_text( angle=90, size=6, vjust=0.5)) 
    
  }
  
  #plot that indicate mean,median and 90% percentile in scatter plot of each variables
  print(marrangeGrob(myplot, ncol = 3, nrow = 2))
  
  #table with min mean median max and SD 
  return(df_descrip)
}

descriptive.table(data_new,"garpts")

############################
### Linear Correlation #####
############################

## Linear correlation for whole dataset
lin.cor <- cor(data_new)

col<- colorRampPalette(c("blue", "white", "red"))(20)
lin.plot <- plot_ly(x=colnames(lin.cor), y=rownames(lin.cor), 
                    z = lin.cor, type = "heatmap", colors = col) %>%
  layout(margin = list(l=120))

lin.plot

#Correlation matrix with significance levels (p-value) - whole data set
correlations <- rcorr(as.matrix(data_new))

#flatting corr matrix
flattenCorrMatrix <- function(cormat, pmat) {
  ut <- upper.tri(cormat)
  data.frame(
    row = rownames(cormat)[row(cormat)[ut]],
    column = rownames(cormat)[col(cormat)[ut]],
    cor  =(cormat)[ut],
    p = pmat[ut]
  )
}

flattenCorrMatrix(correlations$r,correlations$P)


### Create correlation table and plot for each process
corr.func <- function(data,x) {
  
  df <- as.data.frame(data %>% dplyr:: select(starts_with(x)))
  
  corr <- cor(df)
  col<- colorRampPalette(c("blue", "white", "red"))(20)
  
  ut <- upper.tri(corr)
  df.ut <- data.frame(
    row = rownames(corr)[row(corr)[ut]],
    column = rownames(corr)[col(corr)[ut]],
    cor  =(corr)[ut]
  )
  print(df.ut[
    with(df.ut, order(cor, decreasing = TRUE)),
  ] )
  
  plot_ly(x=colnames(corr), y=rownames(corr), 
          z = corr, type = "heatmap", colors = col)
}

corr.func(data_new,"garpts")

################################################# IDENTIFICATION ofPOTENTIAL BUILDS LEADING TO PERFORMAN ANOMALIES #####################################
##############################
# 1st TECHNIQUE: Raw Data + IsolationForest #
##############################
rstudioapi::showDialog(title = "Select sheet",
                       message = "If the number of observation(rows) in data_new 
                       is less than 256, you must enter the number of rows in the following code
                       <b>isolationForest$new(sample_size = number of observation)</b>,
                       in line 265")
iforest = isolationForest$new()
iforest$fit(data_new) # fits an isolation forest for the dataframe

scores_rawdata <- data_new %>%
  iforest$predict() #%>%
#arrange(desc(anomaly_score))
scores_rawdata$build <- rownames(data_new) 
scores_rawdata2 <- arrange(scores_rawdata,desc(anomaly_score))

plot(density(scores_rawdata$anomaly_score), main = "", xlab = "Anomaly score" )
abline(v=0.62, col="red", lty=2, lwd=3)

indices_rawdata <- scores_rawdata[which(scores_rawdata$anomaly_score > 0.61)]
pred_anomalies_rawdata <- data_new[indices_rawdata$id, ]
pred_normal_rawdata <- data_new[-indices_rawdata$id, ]

#print(pred_anomalies_rawdata)

anomalyplot <- plot_ly(scores_rawdata2,x=scores_rawdata2$id, y=scores_rawdata2$anomaly_score, 
                       color = scores_rawdata2$anomaly_score ,
                       type = "scatter",size = scores_rawdata2$anomaly_score,
                       text = paste("Build :", scores_rawdata2$build ,
                                    "AnomalyScore:",round(scores_rawdata2$anomaly_score,2),
                                    sep = "\n"))  %>%
layout(margin = list(l=120))

anomalyplot

umap.pr <- data_new %>%
  scale() %>%
  uwot::umap() %>%
  setNames(c("V1", "V2")) %>%
  as_tibble() %>%
  rowid_to_column() %>%
  left_join(scores_rawdata2, by = c("rowid" = "id"))

decScore <- arrange(umap.pr,desc(scores_rawdata2$anomaly_score))
print(decScore) # the out put is the table that arrange build base on anomaly score
# in decreasing manner

# umap, plot anomaly by transformed data into Uniform Manifold Approximation 
#and Projection (UMAP) which is an algorithm for dimensional reduction proposed 
# and combine it with anomaly score from isolationforest function  
umap.plot <-  ggplot(data = umap.pr,aes(V1, V2)) +
  geom_point(aes(size =scores_rawdata2$anomaly_score))


subplot(anomalyplot,umap.plot)

#########################
# Raw Data + Clustering #
#########################
data_scaled <- scale(data_new)


##FIND OPTIMAL NUMBER FOR K
#1.Elbow mwthod
Elbow_hcl<-fviz_nbclust(data_scaled , hcut, method = "wss") +
  geom_vline(xintercept = 4, linetype = 2) + # add line for better visualisation
  labs(subtitle = "Hierarchical clustering") # add subtitle
#suggest 4 cluster
Elbow_kmean<-fviz_nbclust(data_scaled , kmeans, method = "wss") +
  geom_vline(xintercept = 3, linetype = 2) + # add line for better visualisation
  labs(subtitle = "k-means clustering")

ggarrange(Elbow_hcl, Elbow_kmean , 
          #labels = c("Hierarchical", "k-means"),
          ncol = 2, nrow = 1)
#2.Silhouette method
sil_h <- fviz_nbclust(data_scaled, hcut, method = "silhouette") +
  labs(subtitle = "Hierarchical clustering")

sil_k <- fviz_nbclust(data_scaled, kmeans, method = "silhouette") +
  labs(subtitle = "k-means clustering")
ggarrange(sil_h, sil_k , 
          #labels = c("Hierarchical", "k-means"),
          ncol = 2, nrow = 1)
# suggest 2 cluster
res.k <- eclust(data_scaled, "kmeans", k = 3, hc_metric = "euclidean",
                graph = FALSE) 

res.hc <- eclust(data_scaled, "hclus", k = 3, hc_metric = "euclidean",
                 graph = FALSE) 

#Silhouette plot for hierarchical clustering
sil_hc <- fviz_silhouette(res.hc, print.summary = TRUE)
#ggplotly(sil_hc)

sil_km <- fviz_silhouette(res.k, print.summary = TRUE)
#ggplotly(sil_km)

ggarrange(sil_hc, sil_km , 
          #labels = c("Hierarchical", "k-means"),
          ncol = 2, nrow = 1)
############################################### 
#THIS PART TAKES TIME;EXECUTE IF you WANT TO EXPLORE
# 3.Gap statistic
#It takes too long to run
#set.seed(123)
#fviz_nbclust(data_scaled, hcut,
 #            nstart = 25,
  #           method = "gap_stat",
   #          nboot = 300 # reduce it for lower computation time (but less precise results)
#) +
 # labs(subtitle = "Gap statistic method")

n_clusters(as.data.frame(data_scaled),
           package = c("easystats", "NbClust", "mclust"),
           standardize = FALSE)
#The choice of 3 clusters is supported by 10 (47.62%) methods out
#of 21 (kl, Ch, Cindex, DB, Duda, Pseudot2, Beale, Ratkowsky, Ball, PtBiserial).

nb <- NbClust(data_scaled, diss=NULL, distance = "euclidean",
              method = "ward.D2", min.nc=2, max.nc=20, 
              index = "silhouette")
#hist(nb$Best.nc[1,], breaks = max(na.omit(nb$Best.nc[1,])))
nb$Best.nc
##########################################################################################################################
## HIERACHICAL CLUSTERING
hcl <- hclust(dist(data_scaled), method = 'ward.D2')
plot(hcl, hang = -1, cex = 0.6)
abline(h = 250, lty = 2, col="red")

plt <- fviz_dend(hcl, type = "phylogenic", k = 3) 
ggplotly(plt, tooltip = c("x","y"))
plt2 <- fviz_dend(hcl, type = "rectangle", k = 3)
ggplotly(plt2)


cut_hcl <- cutree(hcl, k=3)

data_hcl <- cbind(cut_hcl,data_scaled)
hcl_1 <- as.data.frame(data_hcl[which(data_hcl[,1]==1),])
hcl_2 <- as.data.frame(data_hcl[which(data_hcl[,1]==2),]) 
hcl_3 <- as.data.frame(data_hcl[which(data_hcl[,1]==3),]) 

dat.raw <- data_hcl[,-1]

clust.centroid = function(i, dat, cut_hcl) {
  ind <- (cut_hcl == i)
  colMeans(dat[ind,])
}

centers <- as.matrix(sapply(unique(cut_hcl), clust.centroid, dat.raw, cut_hcl))
centers <- t(centers)
center <- centers[cut_hcl,]

#Function to find 3 build with larger distance from 
#center of each cluster
cluster_anomaly <- function(dat,i) {
  myplot <- list()
  cl <- subset(dat, cut_hcl==i)
  #hcl_1 <- as.data.frame(data_hcl[which(data_hcl[,1]==1),]) #360
  
  distances <- sqrt(rowSums((cl - center[nrow(cl),])^2))
  # pick top 5 largest distances
  # who are outliers
  print(paste0("Potential Abnormal Builds in Cluster: ",  i))
  print(sort(distances, decreasing = TRUE)[1:5])
  
  par(mfrow=c(1,3))
  myplot[[i]] <- boxplot(distances, main ="Distances of builds from the center of the cluster", col ="orange")
  print(myplot)
}



for(i in 1:3){cluster_anomaly(dat.raw,i)}

#############################################
#2st TECHNIQUE:PRINCIPAL COMPONENT(PCA) #
#############################################
pca <- prcomp(data_new, center = TRUE, scale. = TRUE)

sum.pca <- summary(pca)
ncomp <- which.max(sum.pca$importance[3,] >=0.80)
df.pca  <- as.data.frame(pca $x[,1:ncomp])

pca.data <- data.frame(Sample=rownames(pca$x),
                       X=pca$x[,1],
                       Y=pca$x[,2])

pca.var <- pca$sdev^2
pca.var.per <- round(pca.var/sum(pca.var)*100, 1)

pca.plot <- ggplot(data=pca.data, aes(x=X, y=Y, label=Sample)) +
  modelr::geom_ref_line(h = 0) +
  modelr::geom_ref_line(v = 0) +
  geom_text(size=3) +
  xlab(paste("PC1 - ", pca.var.per[1], "%", sep="")) +
  ylab(paste("PC2 - ", pca.var.per[2], "%", sep="")) +
  ggtitle("My PCA Graph")

pca.plot

##########################
# PCA + IsolationForest #
#########################
rstudioapi::showDialog(title = "Select sheet",
                       message = "If the number of observation(rows) in data_new 
                       is less than 256, you must enter the number of rows in the following code
                       <b>isolationForest$new(sample_size = number of observation)</b>,
                       in line 468")
iforest = isolationForest$new()
iforest$fit(df.pca) # fits an isolation forest for the dataframe

scores_pca <- df.pca %>%
  iforest$predict() #%>%
#arrange(desc(anomaly_score))
plot(density(scores_pca$anomaly_score), main = "", xlab = "Anomaly scores for Cleaned Data+PCA" )
abline(v=0.61, col="red", lty=2, lwd=3)

indices_pca <- scores_pca[which(scores_pca$anomaly_score > 0.61)]
pred_anomalies_pca <- data_new[indices_pca$id, ]
pred_normal_pca <- data_new[-indices_pca$id, ]



# VISYALIZATION ISOLATIONFOREST
scores_pca$build <- rownames(df.pca) 
scores2_pca <- arrange(scores_pca,desc(anomaly_score))

anomalyplot_pca <- plot_ly(scores2_pca,x=scores2_pca$id, y=scores2_pca$anomaly_score, 
                       color = scores2_pca$anomaly_score ,
                       type = "scatter",size = scores2_pca$anomaly_score,
                       text = paste("Build :", scores2_pca$build ,
                                    "AnomalyScore:",round(scores2_pca$anomaly_score,2),
                                    sep = "\n"))  %>%
  layout(margin = list(l=120))


umap.pr_pca <- df.pca %>%
  scale() %>%
  uwot::umap() %>%
  setNames(c("V1", "V2")) %>%
  as_tibble() %>%
  rowid_to_column() %>%
  left_join(scores2_pca, by = c("rowid" = "id"))

decScore_pca <- arrange(umap.pr_pca,desc(scores2_pca$anomaly_score))
print(decScore_pca) # the out put is the table that arrange build base on anomaly score
# in decreasing manner

# umap plot anomaly by transformed data into Uniform Manifold Approximation 
#and Projection (UMAP) which is an algorithm for dimensional reduction proposed 
# and combine it with anomaly score from isolationforest function  
umap.plot_pca <-  ggplot(data = umap.pr_pca,aes(V1, V2)) +
  geom_point(aes(size =scores2_pca$anomaly_score))


subplot(anomalyplot_pca,umap.plot_pca)
####################################################################################################################################################################################
#############################
## PCA + Clustering Analysis#
#############################
df.pca_scaled <- scale(df.pca)

## Find the optimal number of cluster
fviz_nbclust(df.pca_scaled , hcut, method = "wss") +
  geom_vline(xintercept = 4, linetype = 2) + # add line for better visualisation
  labs(subtitle = "Elbow method")

#n_clusters(as.data.frame(df.pca_scaled),
 #          package = c("easystats", "NbClust", "mclust"),
  #         standardize = FALSE)
#2.Silhouette method
fviz_nbclust(df.pca_scaled, hcut, method = "silhouette") +
  labs(subtitle = "Silhouette method")
# suggest 2 cluster
########################################################################
## HIERACHICAL CLUSTERING
hcl_pca <- hclust(dist(df.pca_scaled), method = 'ward.D2')
plot(hcl_pca)
cut_hcl_pca <- cutree(hcl_pca, k=3)

data_hcl_pca <- as.data.frame(cbind(cut_hcl_pca,df.pca_scaled))
hcl_1_pca <- data_hcl_pca[which(data_hcl_pca[,1]==1),]
hcl_2_pca <- data_hcl_pca[which(data_hcl_pca[,1]==2),] 
hcl_3_pca <- data_hcl_pca[which(data_hcl_pca[,1]==3),] 

dat.pca <- data_hcl_pca[,-1]

clust.centroid = function(i, dat, cut_hcl_pca) {
  ind <- (cut_hcl_pca == i)
  colMeans(dat[ind,])
}

centers <- as.matrix(sapply(unique(cut_hcl_pca), clust.centroid, dat.pca, cut_hcl_pca))
centers <- t(centers)
center <- centers[cut_hcl_pca,]

#Function to find 5 build with larger distance from 
#center of each cluster
cluster_anomaly <- function(dat,i) {
  myplot <- list()
  cl <- subset(dat, cut_hcl_pca==i)
  distances <- sqrt(rowSums((cl - center[nrow(cl),])^2))
  # pick top 5 largest distances
  print(paste0("Potential Abnormal Builds in Cluster: ",  i))
  print(sort(distances, decreasing = TRUE)[1:5])
  par(mfrow=c(1,3))
  myplot[[i]] <- boxplot(distances, main ="Distances of builds from the center of the cluster", col ="orange")
  print(myplot)
}

#The output is boxplot and five build with highest distance 
#in each cluster.
for(i in 1:3){cluster_anomaly(dat.pca,i)}

##################################
# Effect of anomaly on multivariate Analysis #
##################################
#Create new dataset without anomalies
#we removed the suspected builds that detected with ISO and Clustering
#in Cleaned Data
anomaly.remove <- c ("7.0.0.0.475","7.0.0.0.522","7.0.0.0.422","7.0.0.0.416","7.0.0.0.410","7.0.0.0.394","7.0.0.0.392",
                    "7.0.0.0.388")

Dat_NoAnomaly <- data_new[!(row.names(data_new) %in% anomaly.remove), ]

##############################
#PRINCIPAL COMPONENT ANALYSIS
##############################
pca_noanomaly <- prcomp(Dat_NoAnomaly, center = TRUE, scale. = TRUE)

sum.pca_noanomaly <- summary(pca_noanomaly)
ncomp_noanomaly <- which.max(sum.pca_noanomaly$importance[3,] >=0.80)
df.pca_noanomaly  <- as.data.frame(pca_noanomaly $x[,1:ncomp_noanomaly])

pca.data_na <- data.frame(Sample=rownames(pca_noanomaly$x),
                       X=pca_noanomaly$x[,1],
                       Y=pca_noanomaly$x[,2])

pca.var_na <- pca_noanomaly$sdev^2
pca.var.per_na <- round(pca.var_na/sum(pca.var_na)*100, 1)

pca.plot_na <- ggplot(data=pca.data_na, aes(x=X, y=Y, label=Sample)) +
  modelr::geom_ref_line(h = 0) +
  modelr::geom_ref_line(v = 0) +
  geom_text(size=3) +
  xlab(paste("PC1 - ", pca.var.per_na[1], "%", sep="")) +
  ylab(paste("PC2 - ", pca.var.per_na[2], "%", sep="")) +
  ggtitle("")

pca.plot_na

##### Cumlative PCA
cumPVE <- qplot(c(1:length(pca$sdev)), cumsum(pca.var.per))+ 
  geom_line() + 
  xlab("Principal Component") + 
  ylab(NULL) + 
  ggtitle("Cumulative Plot for Cleaned Data") +
  ylim(0,100)
cumPVE

cumPVE_na <- qplot(c(1:length(pca_noanomaly$sdev)), cumsum(pca.var.per_na))+ 
  geom_line() + 
  xlab("Principal Component") + 
  ylab(NULL) + 
  ggtitle("Cumulative Plot for Cleaned Data without anomaly") +
  ylim(0,100)

cumPVE_na

#Ranking variables in PC1
#Cleaned Data
loading_scores <- pca$rotation[,1]
var_scores <- abs(loading_scores) ## get the magnitudes
var_score_ranked <- sort(var_scores, decreasing=T)
top_20_var <- names(var_score_ranked[1:20])
top_20_var

#Loadein scores for Data without anomaly in PC1
loading_scores_na <- pca_noanomaly$rotation[,1]
var_scores_na <- abs(loading_scores_na) ## get the magnitudes
var_score_ranked_na <- sort(var_scores_na,decreasing=T)
top_20_var_na <- names(var_score_ranked_na[1:20])
top_20_var_na
#CD_pc_na <- loading_scores_na[(row.names(loading_scores_na) %in% top_20_var_na), ]

#write_xlsx(CD_pc_na,"C://Users//RISE//Desktop//Sara//Master thesis//FINAL//CD_pc_na.xlsx")

#CD_pc <- loading_scores[(row.names(loading_scores) %in% top_20_var), ]
#write_xlsx(CD_pc,"C://Users//RISE//Desktop//Sara//Master thesis//FINAL//CD_pc.xlsx")

#########################################
##MULTIVARIATE LINEAR REGRESSION ANALYSIS
#########################################
