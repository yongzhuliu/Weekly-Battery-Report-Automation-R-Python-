##############
# unzip file #
##############
library(readxl)
library(tidyverse)
library(lubridate)

date <- 20260107
file_name <- "D:\\研究\\weekly"

zipfile <- file.path(Sys.getenv("USERPROFILE"), 
                     "Downloads", 
                     paste0(date,"_cs_newexp.zip"))
unzip(zipfile, exdir = file_name)

#############
# load data #
#############
setwd(file.path(file_name,paste0(date,"_cs_newexp")))
loadfunction <- function(i,j){
  
  cands <- 59:63
  fns <- paste0("240076-", i, "-", j, "-28185739",cands,".xlsx")
  idx <- which(file.exists(fns))[1]
  if (is.na(idx)) stop("Cannot find any file among：", paste(fns, collapse = ", "))
  N <- fns[idx]
  
  TS <- "step"
  part.T <- as.data.frame(read_excel(N, sheet = TS))
  
  S <- "cycle"
  part.NS <- as.data.frame(read_excel(N, sheet = S))
  
  cp_idx <- part.T$工步類型 == "恆流放電"
  capacity <- part.T$`容量(Ah)`[cp_idx]
  cycle <- length(capacity)
  
  name <- paste0("battery",i,j)
  assign(name, part.NS, envir = parent.frame())
  name2 <- paste0("capacity",i,j)
  assign(name2, capacity, envir = parent.frame())
  nameC <- paste0("cycle",i,j)
  assign(nameC, cycle, envir = parent.frame())
  
  statis <- paste0("statis",i,j)
  assign(statis, part.T, envir = parent.frame())
}


for(i in 1:4){
  for(j in 1:8){
    loadfunction(i,j) 
    cat("Battery",i,"-",j," Done\n")
  }
}


###################
# generate report #
###################
dir.create(file.path(file_name,paste0(date,"_cs_newexp"),"report"), 
           showWarnings = FALSE, recursive = TRUE) 
setwd(file.path(file_name,paste0(date,"_cs_newexp"),"report")) 

i <- 1:4
ja <- seq(1,8,by=2)
jb <- seq(2,8,by=2)

# plan A
idx <- t(outer(i, ja, function(ii, jj) 10*ii + jj))
PlanA_cycle <- sapply(idx, function(k) {
  x <- get(paste0("cycle", k), envir = .GlobalEnv)
  x
})
PlanA_capacity <- sapply(idx, function(k) {
  x <- get(paste0("capacity", k), envir = .GlobalEnv)
  x
})
PlanA_standcapacity <- sapply(1:16,function(j){
  PlanA_capacity[[j]]/PlanA_capacity[[j]][1]
})
battery_name_planA <- paste0("Battery ",idx%/%10,"-",idx%%10)

# plan B
idx <- t(outer(i, jb, function(ii, jj) 10*ii + jj))
PlanB_cycle <- sapply(idx, function(k) {
  x <- get(paste0("cycle", k), envir = .GlobalEnv)
  x
})
PlanB_capacity <- sapply(idx, function(k) {
  x <- get(paste0("capacity", k), envir = .GlobalEnv)
  x
})
PlanB_standcapacity <- sapply(1:16,function(j){
  PlanB_capacity[[j]]/PlanB_capacity[[j]][1]
})
battery_name_planB <- paste0("Battery ",idx%/%10,"-",idx%%10)


PlanA_dat_plot <- list(standcapacity = PlanA_standcapacity,
                       cycle = PlanA_cycle, 
                       Battery_name = battery_name_planA)
PlanB_dat_plot <- list(standcapacity = PlanB_standcapacity,
                       cycle = PlanB_cycle, 
                       Battery_name = battery_name_planB)

xlim_num <- c(min(PlanA_cycle),
              min(PlanB_cycle))
cap_range <- list(range(sapply(PlanA_dat_plot$standcapacity, function(x) range(x[1:xlim_num[1]])))+c(-1,1)*0.01,
                  range(sapply(PlanB_dat_plot$standcapacity, function(x) range(x[1:xlim_num[2]])))+c(-1,1)*0.01)

save.image(file = paste0(date,"_cs_newexp.Rdata"))

png("PlanA_capacity.png", 
    width = 8*1.5, height = 8*1.5, units = "in",res = 600)
col_ss = rep(1:3,each = 20)
par(mfrow = c(4,4))
for(i in 1:16){
  plot(1:xlim_num[1], PlanA_dat_plot$standcapacity[[i]][1:xlim_num[1]], col = col_ss, ylim = cap_range[[1]],xlim = c(0,xlim_num[1]),
       pch = 16, xlab = "cycle", ylab = "Standard Capacity", main = PlanA_dat_plot$Battery_name[i],
       cex.axis = 1.8, cex.lab = 1.8, cex.main = 1.8,cex=1.5)
  abline(v=28,col="gray")
}
dev.off()


png("PlanB_capacity.png", 
    width = 8*1.5, height = 8*1.5, units = "in",res = 600)
col_ss = rep(1:3,times = 20)
par(mfrow = c(4,4))
for(i in 1:16){
  plot(1:xlim_num[2], PlanB_dat_plot$standcapacity[[i]][1:xlim_num[2]], col = col_ss, ylim = cap_range[[2]],xlim = c(0,xlim_num[2]),
       pch = 16, xlab = "cycle", ylab = "Standard Capacity", main = PlanB_dat_plot$Battery_name[i],
       cex.axis = 1.8, cex.lab = 1.8, cex.main = 1.8,cex=1.5)
  abline(v=31,col="gray")
}
dev.off()


