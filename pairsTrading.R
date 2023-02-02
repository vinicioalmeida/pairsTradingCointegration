remove(list=ls())
setwd("/Users/...")

library(readxl)
library(writexl)

cotas = read_excel("...", sheet="...", col_names = TRUE)
cotas = as.matrix(cotas)
cotas = cotas[, 2:ncol(cotas)]
n = ncol(cotas) - 1
r = nrow(cotas)
npares = factorial(n) / (factorial(n - 2) * 2)    # Cn,p=(n!)/[(n-p)!*p!]
pares = matrix(nrow = npares + 1, ncol = 14) 
colnames(pares) = c("N",	"Acao1",	"Acao2",	"PP1",	"PP2",	"PPR",	"Alfa",	"Beta",	"Preco1",	"Preco2",	"Max",	"Min",	"DesvioP",	"DesvioAb")
z = 0
for (i in 1:(n - 1)){
  y = as.numeric(cotas[, i])
  for (j in (i + 1):n){
    x = as.numeric(cotas[, j])  
    z = z + 1
    pares[z, 1] = z
    pares[z, 2] = colnames(cotas)[i]
    pares[z, 3] = colnames(cotas)[j]
    pares[z, 4] = PP.test(y)$p.value
    pares[z, 5] = PP.test(x)$p.value
    ols = lm(y ~ x)
    pares[z, 6] = PP.test(ols$residuals)$p.value
    pares[z, 7] = as.numeric(ols$coefficients[1])
    pares[z, 8] = as.numeric(ols$coefficients[2])
    pares[z, 9] = as.numeric(cotas[r, i])
    pares[z, 10] = as.numeric(cotas[r, j])
    pares[z, 11] = max(ols$residuals)
    pares[z, 12] = min(ols$residuals)
    pares[z, 13] = sd(ols$residuals)
    pares[z, 14] = as.numeric(cotas[r, i]) - as.numeric(ols$coefficients[1]) - as.numeric(ols$coefficients[2]) * as.numeric(cotas[r, j])
    
  }
}

paresx = as.data.frame(pares)
write_xlsx(paresx, "C:\\Users\\...")
paresx = as.data.frame(pares)
paresx = as.data.frame(pares)
write.table(paresx, file = 'pares.csv', sep = ';', dec = '.', row.names = FALSE)