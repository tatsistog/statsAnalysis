# clear
cat("\014")
rm(list = ls())

# libraries
library(readxl) #
library(stats) #
library(sur) #
library(e1071) #
library(ggplot2)
library(dplyr) #
library(car) #
library(onewaytests) #
library(plotrix) #
library(DescTools) #
library(agricolae) #

# Deleting descriptives.csv
fn <- "output/descriptives.csv"
#Check its existence
if (file.exists(fn)) 
  #Delete file if it exists
  file.remove(fn)

# filepath
path <- "input/Statistics_New.xlsx"
mydata <- read_excel(path)
column_names <- colnames(mydata)

df <- data.frame()
descriptives_path <- 'output/descriptives.csv'

groups <- c()
vals <- c()

for (i in 1:ncol(mydata)){
  
  sequence_as_title <- data.frame(column_names[i], "", "")
  write.table(sequence_as_title, file = descriptives_path ,append = TRUE, row.names = FALSE, col.names = FALSE)
  
  second_matrix <- data.frame()
  curr_col = mydata[[i]]
  to_be_deleted = c()
  for (j in length(curr_col):1){
    
    if (is.na(curr_col[j])){
      to_be_deleted <- c(j, to_be_deleted)
    }
    else{
      break
    }
  }
  if (!length(to_be_deleted)==0){
    curr_col <- curr_col[-c(to_be_deleted)]
  }
  
  # Constructing the first matrix
  total <- length(curr_col)
  valid <- sum(!is.na(curr_col))
  missing <- total - valid
  
  valid_perc <- valid/total
  missing_perc <- missing/total
  
  de <- data.frame(valid, valid_perc, missing, missing_perc, total)
  names(de) <- c("Valid", "Valid_Percentage", "Missing", "Missing_Percentage", "Total")
  df <- rbind(df, de)
  
  # Constructing the second matrix
  av <- mean(curr_col, na.rm = TRUE)
  std_error <- plotrix::std.error(curr_col)

  s <- var(curr_col, na.rm = TRUE)
  st_dev <- sqrt(s)
  confidence_int <- t.test(curr_col)
  
  conf_left <- confidence_int$conf.int[1]
  conf_right <- confidence_int$conf.int[2]
  
  med <- median(curr_col)
  col_min <- min(curr_col)
  col_max <- max(curr_col)
  col_range <- col_max - col_min
  
  av_trimmed = mean(sort(curr_col), trim =  0.05, na.rm = TRUE)
  curr_col_iqr <- IQR(curr_col, na.rm = TRUE, type = 6)
  curr_col_skewness <- skewness(curr_col)
  skewness_error <- se.skew(curr_col)
  curr_col_kurtosis <- kurtosis(curr_col)
  kurtosis_error <- sqrt(4*(valid^2 -1 ) * (skewness_error^2) / ((valid -3  )*(valid+5)))
  
  de <- data.frame(c(av, conf_left, conf_right, av_trimmed, med, s, st_dev, col_min, col_max, col_range, curr_col_iqr, curr_col_skewness, curr_col_kurtosis), c(std_error, "", "", "","","","","","","","", skewness_error, kurtosis_error))
  names(de) <- data.frame("Statistic", "Std. Error")
  
  write.table(data.frame("","Statistic","Std. Error"), file = descriptives_path ,append = TRUE, row.names = FALSE, col.names = FALSE)
  second_matrix <- rbind(second_matrix, de)
  row.names(second_matrix) <- c("Mean", "95% conf. int lower", "95% conf. int upper", "5% trimmed mean", "Median", "Variance", "Std Deviation", "Minimum", "Maximum", "Range", "Interquartile Range", "Skewness", "Kurtosis")
  write.table(second_matrix, file = descriptives_path ,append = TRUE, row.names = TRUE, col.names = FALSE)
  write.table(data.frame("","",""), file = descriptives_path ,append = TRUE, row.names = FALSE, col.names = FALSE)
  
  vals <- c(vals, curr_col)
  groups <- c(groups, rep(column_names[i],length(curr_col)))
  
}

row.names(df) <- colnames(mydata)
data <- as.data.frame(df, row.names = colnames(mydata))
write.csv(data, file = 'output/Valid_missing_frames.csv')

png(filename = 'output/boxplots.png', height = 600, width = 1000)
boxplot(as.matrix(mydata), xlab = "Categories_new", ylab = "Measure", main = "Box Plots")
dev.off()

levTest <- leveneTest(vals, groups, center = "mean")
df <- data.frame(levTest$`F value`[1], levTest$Df[1], levTest$Df[2], levTest$`Pr(>F)`[1])
names(df) <- c("Levene Statistic", "df1", "df2", "Sig.")
write.csv(df, file = 'output/leveneTest.csv')

cur_data <- data.frame(vals, groups)
names(cur_data) <- c("values", "groups")

anovaTest <- anova(lm(vals~groups))
df <- data.frame(c(anovaTest$`Sum Sq`[1], anovaTest$`Sum Sq`[2], anovaTest$`Sum Sq`[1]+anovaTest$`Sum Sq`[2]), 
                 c(anovaTest$Df[1], anovaTest$Df[2], anovaTest$Df[1]+ anovaTest$Df[2]), 
                 c(anovaTest$`Mean Sq`[1], anovaTest$`Mean Sq`[2], ""), 
                 c(anovaTest$`F value`[1],"" ,""), c(anovaTest$`Pr(>F)`[1],"",""))
names(df) <- c("Sum of Squares", "df", "Mean Square", "F", "Sig.")
row.names(df) <- c("Between Groups", "Within Groups", "Total")
write.csv(df, file = 'output/anovaTest.csv')

welchtest <- welch.test(values~groups, cur_data)
row1 <- c(welchtest$statistic, welchtest$parameter[1], welchtest$parameter[2], welchtest$p.value)

cur_data$groups <- factor(cur_data$groups)
bftest <- bf.test(values~groups, cur_data)
row2 <- c(bftest$statistic, bftest$parameter[1], bftest$parameter[2], bftest$p.value)

robust_matrix <- data.frame(rbind(row1,row2))
names(robust_matrix) <- c("Statistc", "df1", "df2", "Sig.")
row.names(robust_matrix) <- c("Welch", "Brown-Forsythe")
write.csv(robust_matrix, 'output/robust_tests.csv')


# Post Hoc Analysis:
myrownames = c()
mean_ij <- c()
left_int <- c()
right_int <- c()
sig <- c()
a1 <- aov(lm(vals ~ groups))
pwc <- PostHocTest(a1, method = "bonferroni", conf.level = 0.95)
pwc <- pwc$groups
mswithin <- anovaTest$`Mean Sq`[2]
std_error_vector <- c()


for (i in 1:ncol(mydata)){
  
  for (j in 1:ncol(mydata)){
    if (i == j){
      next
    }
    cur_vals <- c(vals[which(cur_data$groups == column_names[i])],  vals[which(cur_data$groups == column_names[j])])
    cur_groups <- c(groups[which(cur_data$groups == column_names[i])],  groups[which(cur_data$groups == column_names[j])])
    mse_within <- sum(cur_vals^2)/(length(cur_vals)-2)
    N1 <- length(which(cur_groups == levels(factor(cur_groups))[1]))
    N2 <- length(which(cur_groups == levels(factor(cur_groups))[2]))
    std_error <- sqrt(mswithin*((1/N1) + (1/N2)))
    std_error_vector <- c(std_error_vector, std_error)
    
    myrownames <- c(myrownames, paste(column_names[i], column_names[j] , sep = "-"))
    name1 <- paste(column_names[i], column_names[j] , sep = "-")
    name2 <- paste(column_names[j], column_names[i] , sep = "-")
    
    find1 <- which(rownames(pwc) == name1)
    find2 <- which(rownames(pwc) == name2)
    
    if (length(find1) == 0){
      
      mean_ij <- c(mean_ij, -pwc[find2,1])
      left_int <- c(left_int, -pwc[find2,3])
      right_int <- c(right_int, -pwc[find2,2])
      sig <- c(sig, pwc[find2, 4])
      
    }else{
      mean_ij <- c(mean_ij, pwc[find1,1])
      left_int <- c(left_int, pwc[find1,2])
      right_int <- c(right_int, pwc[find1,3])
      sig <- c(sig, pwc[find1, 4])
    }
  }
}

post_hoc_table = data.frame(mean_ij, std_error_vector, sig, left_int, right_int)
names(post_hoc_table) <- c("Mean Difference", "Std.Error","Sig", "95% conf.int lower Bound", "95% conf.int Upper Bound")
row.names(post_hoc_table) <- myrownames
write.csv(post_hoc_table, 'output/Bonferroni_post_hoc_tests.csv')


duncan_test <- duncan.test(a1, 'groups',MSerror = mswithin, alpha = 0.05, console = TRUE, group = FALSE)
df <- data.frame(duncan_test$means$r, duncan_test$means$vals)
names(df) <- c("N",  "vals")
row.names(df) <- c(rownames(duncan_test$means))
write.csv(df, 'output/duncan_test.csv')
write.csv(duncan_test$comparison, 'output/duncan_test_sig.csv')

kruskal_test <- kruskal.test(values~groups, data = cur_data)
kt2 <- kruskal(vals, groups, alpha = 0.05, p.adj = "bonferroni", group = TRUE)
df <- data.frame(c(kruskal_test$statistic, kruskal_test$parameter, kruskal_test$p.value))
row.names(df)[3] <- "Asymp. Sig"
names(df) <- c("Measure")
write.csv(df, 'output/Kruskal_Wallis_Test_Statistics.csv')

ranks_table <- data.frame(kt2$means$r, kt2$means$rank)
row.names(ranks_table) <- rownames(kt2$means)
names(ranks_table) <- c("N", "Mean Rank")
write.csv(ranks_table, 'output/Kruskal_Wallis_Test_Mean_Ranks.csv')