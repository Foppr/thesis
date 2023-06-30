anova_summary_to_dataframe = function(anova){
  
  summ = summary(anova)
  
  df_summary = data.frame(Source = rownames(summ[[1]]),
                          DF = summ[[1]]$Df,
                          SS = summ[[1]]$`Sum Sq`,
                          MS = summ[[1]]$`Mean Sq`,
                          Fvalue = summ[[1]]$`F value`,
                          PRvalue = summ[[1]]$`Pr(>F)`)
  
  df_summary
  
}

lm_summary_to_dataframe = function(lm){
  library(dplyr)
  lm_summary = summary(lm)
  df_lm = as.data.frame(coef(lm_summary))
  df_lm$coefficients <- rownames(df_lm)
  df_lm <- df_lm[, c("coefficients", "Estimate", "Std. Error", "t value", "Pr(>|t|)")]
  rownames(df_lm) <- NULL
  df_lm
  
}

export_results_1 = function(talk_id){
  library(readxl)
  library(xlsx)
  library(openxlsx)
  library(caret)
  library(dplyr)
  
  pathway = '' # add pathway here
  id = as.character(talk_id)
  
  pathway_id = paste(pathway, id, '/', sep='')
  pathway_id_R = paste(pathway_id, 'R_', id, '.xlsx', sep='')
  
  df = read_excel(pathway_id_R)
  df = na.omit(df)
  
  # Complete model
  agg = aggregate(alignment ~ actual_sense, data = df, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod = aov(alignment ~ actual_sense, data=df)
  agg_l1 = aggregate(alignment ~ sense_l1, data = df, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_l1 = aov(alignment ~ sense_l1, data=df)
  
  # Subset 1 variable
  df_REP = subset(df, repetition == 'no')
  agg_REP = aggregate(alignment ~ actual_sense, data = df_REP, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP = aov(alignment ~ actual_sense, data=df_REP)
  agg_REP_l1 = aggregate(alignment ~ sense_l1, data = df_REP, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_l1 = aov(alignment ~ sense_l1, data=df_REP)
  
  df_UNR = subset(df, relation_type_expected %in% c('Related', 'Hypophora'))
  agg_UNR = aggregate(alignment ~ actual_sense, data = df_UNR, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_UNR = aov(alignment ~ actual_sense, data=df_UNR)
  agg_UNR_l1 = aggregate(alignment ~ sense_l1, data = df_UNR, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_UNR_l1 = aov(alignment ~ sense_l1, data=df_UNR)
   
  df_CLO = subset(df, question_type == 'open')
  agg_CLO = aggregate(alignment ~ actual_sense, data = df_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_CLO = aov(alignment ~ actual_sense, data=df_CLO)
  agg_CLO_l1 = aggregate(alignment ~ sense_l1, data = df_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_CLO_l1 = aov(alignment ~ sense_l1, data=df_CLO)
  
  # Subset 2 variables
  df_REP_UNR = subset(df, repetition == 'no' & relation_type_expected %in% c('Related', 'Hypophora'))
  agg_REP_UNR = aggregate(alignment ~ actual_sense, data = df_REP_UNR, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_UNR = aov(alignment ~ actual_sense, data=df_REP_UNR)
  agg_REP_UNR_l1 = aggregate(alignment ~ sense_l1, data = df_REP_UNR, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_UNR_l1 = aov(alignment ~ sense_l1, data=df_REP_UNR)
  
  df_REP_CLO = subset(df, repetition == 'no' & question_type == 'open')
  agg_REP_CLO = aggregate(alignment ~ actual_sense, data = df_REP_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_CLO = aov(alignment ~ actual_sense, data=df_REP_CLO)
  agg_REP_CLO_l1 = aggregate(alignment ~ sense_l1, data = df_REP_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_CLO_l1 = aov(alignment ~ sense_l1, data=df_REP_CLO)
  
  df_UNR_CLO = subset(df, relation_type_expected %in% c('Related', 'Hypophora') & question_type == 'open')
  agg_UNR_CLO = aggregate(alignment ~ actual_sense, data = df_UNR_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_UNR_CLO = aov(alignment ~ actual_sense, data=df_UNR_CLO)
  agg_UNR_CLO_l1 = aggregate(alignment ~ sense_l1, data = df_UNR_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_UNR_CLO_l1 = aov(alignment ~ sense_l1, data=df_UNR_CLO)
  
  # Subset all 3 variables
  df_REP_UNR_CLO = subset(df, repetition == 'no' & relation_type_expected %in% c('Related', 'Hypophora') & question_type == 'open')
  agg_REP_UNR_CLO = aggregate(alignment ~ actual_sense, data = df_REP_UNR_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_UNR_CLO = aov(alignment ~ actual_sense, data=df_REP_UNR_CLO)
  agg_REP_UNR_CLO_l1 = aggregate(alignment ~ sense_l1, data = df_REP_UNR_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_UNR_CLO_l1 = aov(alignment ~ sense_l1, data=df_REP_UNR_CLO)
  
  # Cross-validation for finding the best model
  cross_validate = train(alignment ~ actual_sense, data = df, method = "lm")
  cross_validate_REP = train(alignment ~ actual_sense, data = df_REP, method = "lm")
  cross_validate_UNR = train(alignment ~ actual_sense, data = df_UNR, method = "lm")
  cross_validate_CLO = train(alignment ~ actual_sense, data = df_CLO, method = "lm")
  cross_validate_REP_UNR = train(alignment ~ actual_sense, data = df_REP_UNR, method = "lm")
  cross_validate_REP_CLO = train(alignment ~ actual_sense, data = df_REP_CLO, method = "lm")
  cross_validate_UNR_CLO = train(alignment ~ actual_sense, data = df_UNR_CLO, method = "lm")
  cross_validate_REP_UNR_CLO = train(alignment ~ actual_sense, data = df_REP_UNR_CLO, method = "lm")
  
  #view summary of k-fold CV   
  df_CV = as.data.frame(cross_validate$results)
  df_CV_REP = as.data.frame = (cross_validate_REP$results)
  df_CV_UNR = as.data.frame = (cross_validate_UNR$results)
  df_CV_CLO = as.data.frame = (cross_validate_CLO$results)
  df_CV_REP_UNR = as.data.frame = (cross_validate_REP_UNR$results)
  df_CV_REP_CLO = as.data.frame = (cross_validate_REP_CLO$results)
  df_CV_UNR_CLO = as.data.frame = (cross_validate_UNR_CLO$results)
  df_CV_REP_UNR_CLO = as.data.frame = (cross_validate_REP_UNR_CLO$results)
  
  df_cross_validation_1 = bind_rows(df_CV, 
                                    df_CV_REP, 
                                    df_CV_UNR, 
                                    df_CV_CLO, 
                                    df_CV_REP_UNR, 
                                    df_CV_REP_CLO, 
                                    df_CV_UNR_CLO, 
                                    df_CV_REP_UNR_CLO)
  
  names = c("df_CV", "df_CV_REP", "df_CV_UNR", "df_CV_CLO", "df_CV_REP_UNR", "df_CV_REP_CLO", "df_CV_UNR_CLO", "df_CV_REP_UNR_CLO")
  df_names = as.data.frame(names)
  df_cross_validation = bind_cols(df_names, df_cross_validation_1)
  
  # Create dataframes of 1) the aggregate summary stats, 2) anova summaries and 3) TukeyHSDs
  # 1)
  df_agg = as.data.frame(agg)
  df_agg_REP = as.data.frame(agg_REP)
  df_agg_UNR = as.data.frame(agg_UNR)
  df_agg_CLO = as.data.frame(agg_CLO)
  df_agg_REP_UNR = as.data.frame(agg_REP_UNR)
  df_agg_REP_CLO = as.data.frame(agg_REP_CLO)
  df_agg_UNR_CLO = as.data.frame(agg_UNR_CLO)
  df_agg_REP_UNR_CLO = as.data.frame(agg_REP_UNR_CLO)
  
  df_agg_l1 = as.data.frame(agg_l1)
  df_agg_REP_l1 = as.data.frame(agg_REP_l1)
  df_agg_UNR_l1 = as.data.frame(agg_UNR_l1)
  df_agg_CLO_l1 = as.data.frame(agg_CLO_l1)
  df_agg_REP_UNR_l1 = as.data.frame(agg_REP_UNR_l1)
  df_agg_REP_CLO_l1 = as.data.frame(agg_REP_CLO_l1)
  df_agg_UNR_CLO_l1 = as.data.frame(agg_UNR_CLO_l1)
  df_agg_REP_UNR_CLO_l1 = as.data.frame(agg_REP_UNR_CLO_l1)
  
  # 2)
  df_sum = anova_summary_to_dataframe(mod)
  df_sum_REP = anova_summary_to_dataframe(mod_REP)
  df_sum_UNR = anova_summary_to_dataframe(mod_UNR)
  df_sum_CLO = anova_summary_to_dataframe(mod_CLO)
  df_sum_REP_UNR = anova_summary_to_dataframe(mod_REP_UNR)
  df_sum_REP_CLO = anova_summary_to_dataframe(mod_REP_CLO)
  df_sum_UNR_CLO = anova_summary_to_dataframe(mod_UNR_CLO)
  df_sum_REP_UNR_CLO = anova_summary_to_dataframe(mod_REP_UNR_CLO)
  
  df_sum_l1 = anova_summary_to_dataframe(mod_l1)
  df_sum_REP_l1 = anova_summary_to_dataframe(mod_REP_l1)
  df_sum_UNR_l1 = anova_summary_to_dataframe(mod_UNR_l1)
  df_sum_CLO_l1 = anova_summary_to_dataframe(mod_CLO_l1)
  df_sum_REP_UNR_l1 = anova_summary_to_dataframe(mod_REP_UNR_l1)
  df_sum_REP_CLO_l1 = anova_summary_to_dataframe(mod_REP_CLO_l1)
  df_sum_UNR_CLO_l1 = anova_summary_to_dataframe(mod_UNR_CLO_l1)
  df_sum_REP_UNR_CLO_l1 = anova_summary_to_dataframe(mod_REP_UNR_CLO_l1)
  
  Tuk = TukeyHSD(mod, conf.level=.95)
  Tuk_REP = TukeyHSD(mod_REP, conf.level=.95)
  Tuk_UNR = TukeyHSD(mod_UNR, conf.level=.95)
  Tuk_CLO = TukeyHSD(mod_CLO, conf.level=.95)
  Tuk_REP_UNR = TukeyHSD(mod_REP_UNR, conf.level=.95)
  Tuk_REP_CLO = TukeyHSD(mod_REP_CLO, conf.level=.95)
  Tuk_UNR_CLO = TukeyHSD(mod_UNR_CLO, conf.level=.95)
  Tuk_REP_UNR_CLO = TukeyHSD(mod_REP_UNR_CLO, conf.level=.95)
  
  Tuk_l1 = TukeyHSD(mod_l1, conf.level=.95)
  Tuk_REP_l1 = TukeyHSD(mod_REP_l1, conf.level=.95)
  Tuk_UNR_l1 = TukeyHSD(mod_UNR_l1, conf.level=.95)
  Tuk_CLO_l1 = TukeyHSD(mod_CLO_l1, conf.level=.95)
  Tuk_REP_UNR_l1 = TukeyHSD(mod_REP_UNR_l1, conf.level=.95)
  Tuk_REP_CLO_l1 = TukeyHSD(mod_REP_CLO_l1, conf.level=.95)
  Tuk_UNR_CLO_l1 = TukeyHSD(mod_UNR_CLO_l1, conf.level=.95)
  Tuk_REP_UNR_CLO_l1 = TukeyHSD(mod_REP_UNR_CLO_l1, conf.level=.95)
  
  # 3)
  # senses
  df_Tuk = as.data.frame(Tuk$actual_sense)
  df_Tuk$comparison <- rownames(df_Tuk)
  df_Tuk <- df_Tuk[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk) <- NULL
  
  df_Tuk_REP = as.data.frame(Tuk_REP$actual_sense)
  df_Tuk_REP$comparison <- rownames(df_Tuk_REP)
  df_Tuk_REP <- df_Tuk_REP[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP) <- NULL
  
  df_Tuk_UNR = as.data.frame(Tuk_UNR$actual_sense)
  df_Tuk_UNR$comparison <- rownames(df_Tuk_UNR)
  df_Tuk_UNR <- df_Tuk_UNR[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_UNR) <- NULL
  
  df_Tuk_CLO = as.data.frame(Tuk_CLO$actual_sense)
  df_Tuk_CLO$comparison <- rownames(df_Tuk_CLO)
  df_Tuk_CLO <- df_Tuk_CLO[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_CLO) <- NULL
  
  df_Tuk_REP_UNR = as.data.frame(Tuk_REP_UNR$actual_sense)
  df_Tuk_REP_UNR$comparison <- rownames(df_Tuk_REP_UNR)
  df_Tuk_REP_UNR <- df_Tuk_REP_UNR[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_UNR) <- NULL
  
  df_Tuk_REP_CLO = as.data.frame(Tuk_REP_CLO$actual_sense)
  df_Tuk_REP_CLO$comparison <- rownames(df_Tuk_REP_CLO)
  df_Tuk_REP_CLO <- df_Tuk_REP_CLO[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_CLO) <- NULL
  
  df_Tuk_UNR_CLO = as.data.frame(Tuk_UNR_CLO$actual_sense)
  df_Tuk_UNR_CLO$comparison <- rownames(df_Tuk_UNR_CLO)
  df_Tuk_UNR_CLO <- df_Tuk_UNR_CLO[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_UNR_CLO) <- NULL
  
  df_Tuk_REP_UNR_CLO = as.data.frame(Tuk_REP_UNR_CLO$actual_sense)
  df_Tuk_REP_UNR_CLO$comparison <- rownames(df_Tuk_REP_UNR_CLO)
  df_Tuk_REP_UNR_CLO <- df_Tuk_REP_UNR_CLO[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_UNR_CLO) <- NULL
  
  # l1
  df_Tuk_l1 = as.data.frame(Tuk_l1$sense_l1)
  df_Tuk_l1$comparison <- rownames(df_Tuk_l1)
  df_Tuk_l1 <- df_Tuk_l1[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_l1) <- NULL
  
  df_Tuk_REP_l1 = as.data.frame(Tuk_REP_l1$sense_l1)
  df_Tuk_REP_l1$comparison <- rownames(df_Tuk_REP_l1)
  df_Tuk_REL_l1 <- df_Tuk_REP_l1[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_l1) <- NULL
  
  df_Tuk_UNR_l1 = as.data.frame(Tuk_UNR_l1$sense_l1)
  df_Tuk_UNR_l1$comparison <- rownames(df_Tuk_UNR_l1)
  df_Tuk_UNR_l1 <- df_Tuk_UNR_l1[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_UNR_l1) <- NULL
  
  df_Tuk_CLO_l1 = as.data.frame(Tuk_CLO_l1$sense_l1)
  df_Tuk_CLO_l1$comparison <- rownames(df_Tuk_CLO_l1)
  df_Tuk_CLO_l1 <- df_Tuk_CLO_l1[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_CLO_l1) <- NULL
  
  df_Tuk_REP_UNR_l1 = as.data.frame(Tuk_REP_UNR__l1$sense_l1)
  df_Tuk_REP_UNR__l1$comparison <- rownames(df_Tuk_REP_UNR__l1)
  df_Tuk_REP_UNR__l1 <- df_Tuk_REP_UNR__l1[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_UNR__l1) <- NULL
  
  df_Tuk_REP_CLO_l1 = as.data.frame(Tuk_REP_CLO_l1$sense_l1)
  df_Tuk_REP_CLO_l1$comparison <- rownames(df_Tuk_REP_CLO_l1)
  df_Tuk_REP_CLO_l1 <- df_Tuk_REP_CLO_l1[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_CLO_l1) <- NULL
  
  df_Tuk_UNR_CLO_l1 = as.data.frame(Tuk_UNR_CLO_l1$sense_l1)
  df_Tuk_UNR_CLO_l1$comparison <- rownames(df_Tuk_UNR_CLO_l1)
  df_Tuk_UNR_CLO_l1 <- df_Tuk_UNR_CLO_l1[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_UNR_CLO_l1) <- NULL
  
  df_Tuk_REP_UNR_CLO_l1 = as.data.frame(Tuk_REP_UNR_CLO_l1$sense_l1)
  df_Tuk_REP_UNR_CLO_l1$comparison <- rownames(df_Tuk_REP_UNR_CLO_l1)
  df_Tuk_REP_UNR_CLO_l1 <- df_Tuk_REP_UNR_CLO_l1[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_UNR_CLO_l1) <- NULL
  
  # Create a list of dataframes
  df_list = list(df_cross_validation = df_cross_validation, 
                 df_sum = df_sum, 
                 df_Tuk = df_Tuk, 
                 df_sum_REP = df_sum_REP, 
                 df_Tuk_REP = df_Tuk_REP, 
                 df_sum_UNR = df_sum_UNR, 
                 df_Tuk_UNR = df_Tuk_UNR, 
                 df_sum_CLO = df_sum_CLO,                
                 df_Tuk_CLO = df_Tuk_CLO, 
                 df_sum_REP_UNR = df_sum_REP_UNR, 
                 df_Tuk_REP_UNR = df_Tuk_REP_UNR, 
                 df_sum_REP_CLO = df_sum_REP_CLO, 
                 df_Tuk_REP_CLO = df_Tuk_REP_CLO, 
                 df_sum_UNR_CLO = df_sum_UNR_CLO, 
                 df_Tuk_UNR_CLO = df_Tuk_UNR_CLO, 
                 df_sum_REP_UNR_CLO = df_sum_REP_UNR_CLO, 
                 df_Tuk_REP_UNR_CLO = df_Tuk_REP_UNR_CLO
  )
  
  # Export every df_1 to its own Excel sheet in one Excel file
  require(openxlsx)
  pathway_id_results_1 = paste(pathway_id, 'results_', id, '_1.xlsx', sep='')
  wb = createWorkbook()
  
  addWorksheet(wb, "cross_validation")
  writeData(wb, "cross_validation", df_cross_validation)
  
  ncol_agg = ncol(df_agg)
  ncol_sum = ncol(df_sum)
  addWorksheet(wb, "df")
  writeData(wb, "df", df_Tuk, startCol = 1, startRow = 1)
  writeData(wb, "df", df_sum, startCol = 7, startRow = 1)
  writeData(wb, "df", df_agg, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "REP")
  writeData(wb, "REP", df_Tuk_REP, startCol = 1, startRow = 1)
  writeData(wb, "REP", df_sum_REP, startCol = 7, startRow = 1)
  writeData(wb, "REP", df_agg_REP, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "UNR")
  writeData(wb, "UNR", df_Tuk_UNR, startCol = 1, startRow = 1)
  writeData(wb, "UNR", df_sum_UNR, startCol = 7, startRow = 1)
  writeData(wb, "UNR", df_agg_UNR, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "CLO")
  writeData(wb, "CLO", df_Tuk_CLO, startCol = 1, startRow = 1)
  writeData(wb, "CLO", df_sum_CLO, startCol = 7, startRow = 1)
  writeData(wb, "CLO", df_agg_CLO, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "REP_UNR")
  writeData(wb, "REP_UNR", df_Tuk_REP_UNR, startCol = 1, startRow = 1)
  writeData(wb, "REP_UNR", df_sum_REP_UNR, startCol = 7, startRow = 1)
  writeData(wb, "REP_UNR", df_agg_REP_UNR, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "REP_CLO")
  writeData(wb, "REP_CLO", df_Tuk_REP_CLO, startCol = 1, startRow = 1)
  writeData(wb, "REP_CLO", df_sum_REP_CLO, startCol = 7, startRow = 1)
  writeData(wb, "REP_CLO", df_agg_REP_CLO, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "UNR_CLO")
  writeData(wb, "UNR_CLO", df_Tuk_UNR_CLO, startCol = 1, startRow = 1)
  writeData(wb, "UNR_CLO", df_sum_UNR_CLO, startCol = 7, startRow = 1)
  writeData(wb, "UNR_CLO", df_agg_UNR_CLO, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "REP_UNR_CLO")
  writeData(wb, "REP_UNR_CLO", df_Tuk_REP_UNR_CLO, startCol = 1, startRow = 1)
  writeData(wb, "REP_UNR_CLO", df_sum_REP_UNR_CLO, startCol = 7, startRow = 1)
  writeData(wb, "REP_UNR_CLO", df_agg_REP_UNR_CLO, startCol = 14, startRow = 1)
  
  saveWorkbook(wb, pathway_id_results_1, overwrite = TRUE)  # Save the workbook to the specified pathway
  
}

export_results_2 = function(talk_id){
  library(readxl)
  library(xlsx)
  library(openxlsx)
  library(caret)
  
  pathway = '' # add pathway here
  id = as.character(talk_id)
  
  pathway_id = paste(pathway, id, '/', sep='')
  pathway_id_R = paste(pathway_id, 'R_', id, '.xlsx', sep='')
  
  df = read_excel(pathway_id_R)
  
  len_df = length(df$q_id)
  
  level = c(rep('Level-1', len_df), rep('Level-2', len_df), rep('Level-3', len_df))
  alignment_levels = c(df$alignment_l1, df$alignment_l2, df$alignment_l3)
  q_id = c(rep(df$q_id, 3))
  relation_type_expected = c(rep(df$relation_type_expected, 3))
  repetition = c(rep(df$repetition, 3))
  question_type = c(rep(df$question_type, 3))
  df_stacked = data.frame(q_id, alignment_levels, level, relation_type_expected, repetition, question_type)
  # new stacked df
  
  # Complete model
  agg = aggregate(alignment_levels ~ level, data = df_stacked, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod = aov(alignment_levels ~ level, data=df_stacked)
  
  # Subset 1 variable
  df_REP = subset(df_stacked, repetition == 'no')
  agg_REP = aggregate(alignment_levels ~ level, data = df_REP, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP = aov(alignment_levels ~ level, data=df_REP)
  
  df_UNR = subset(df_stacked, relation_type_expected %in% c('Related', 'Hypophora'))
  agg_UNR = aggregate(alignment_levels ~ level, data = df_UNR, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_UNR = aov(alignment_levels ~ level, data=df_UNR)
  
  df_CLO = subset(df_stacked, question_type == 'open')
  agg_CLO = aggregate(alignment_levels ~ level, data = df_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_CLO = aov(alignment_levels ~ level, data=df_CLO)
  
  # Subset 2 variables
  df_REP_UNR = subset(df_stacked, repetition == 'no' & relation_type_expected %in% c('Related', 'Hypophora'))
  agg_REP_UNR = aggregate(alignment_levels ~ level, data = df_REP_UNR, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_UNR = aov(alignment_levels ~ level, data=df_REP_UNR)
  
  df_REP_CLO = subset(df_stacked, repetition == 'no' & question_type == 'open')
  agg_REP_CLO = aggregate(alignment_levels ~ level, data = df_REP_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_CLO = aov(alignment_levels ~ level, data=df_REP_CLO)
  
  df_UNR_CLO = subset(df_stacked, relation_type_expected %in% c('Related', 'Hypophora') & question_type == 'open')
  agg_UNR_CLO = aggregate(alignment_levels ~ level, data = df_UNR_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_UNR_CLO = aov(alignment_levels ~ level, data=df_UNR_CLO)
  
  # Subset all 3 variables
  df_REP_UNR_CLO = subset(df_stacked, repetition == 'no' & relation_type_expected %in% c('Related', 'Hypophora') & question_type == 'open')
  agg_REP_UNR_CLO = aggregate(alignment_levels ~ level, data = df_REP_UNR_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_UNR_CLO = aov(alignment_levels ~ level, data=df_REP_UNR_CLO)
  
  # Cross-validation for finding the best model
  cross_validate = train(alignment_levels ~ level, data = df_stacked, method = "lm")
  cross_validate_REP = train(alignment_levels ~ level, data = df_REP, method = "lm")
  cross_validate_UNR = train(alignment_levels ~ level, data = df_UNR, method = "lm")
  cross_validate_CLO = train(alignment_levels ~ level, data = df_CLO, method = "lm")
  cross_validate_REP_UNR = train(alignment_levels ~ level, data = df_REP_UNR, method = "lm")
  cross_validate_REP_CLO = train(alignment_levels ~ level, data = df_REP_CLO, method = "lm")
  cross_validate_UNR_CLO = train(alignment_levels ~ level, data = df_UNR_CLO, method = "lm")
  cross_validate_REP_UNR_CLO = train(alignment_levels ~ level, data = df_REP_UNR_CLO, method = "lm")
  
  #view summary of k-fold CV   
  df_CV = as.data.frame(cross_validate$results)
  df_CV_REP = as.data.frame = (cross_validate_REP$results)
  df_CV_UNR = as.data.frame = (cross_validate_UNR$results)
  df_CV_CLO = as.data.frame = (cross_validate_CLO$results)
  df_CV_REP_UNR = as.data.frame = (cross_validate_REP_UNR$results)
  df_CV_REP_CLO = as.data.frame = (cross_validate_REP_CLO$results)
  df_CV_UNR_CLO = as.data.frame = (cross_validate_UNR_CLO$results)
  df_CV_REP_UNR_CLO = as.data.frame = (cross_validate_REP_UNR_CLO$results)
  
  df_cross_validation_1 = bind_rows(df_CV, 
                                    df_CV_REP, 
                                    df_CV_UNR, 
                                    df_CV_CLO, 
                                    df_CV_REP_UNR, 
                                    df_CV_REP_CLO, 
                                    df_CV_UNR_CLO, 
                                    df_CV_REP_UNR_CLO)
  
  names = c("df_CV", "df_CV_REP", "df_CV_UNR", "df_CV_CLO", "df_CV_REP_UNR", "df_CV_REP_CLO", "df_CV_UNR_CLO", "df_CV_REP_UNR_CLO")
  df_names = as.data.frame(names)
  df_cross_validation = bind_cols(df_names, df_cross_validation_1)
  
  # Create dataframes of 1) the aggregate summary stats, 2) anova summaries and 3) TukeyHSDs
  # 1)
  df_agg = as.data.frame(agg)
  df_agg_REP = as.data.frame(agg_REP)
  df_agg_UNR = as.data.frame(agg_UNR)
  df_agg_CLO = as.data.frame(agg_CLO)
  df_agg_REP_UNR = as.data.frame(agg_REP_UNR)
  df_agg_REP_CLO = as.data.frame(agg_REP_CLO)
  df_agg_UNR_CLO = as.data.frame(agg_UNR_CLO)
  df_agg_REP_UNR_CLO = as.data.frame(agg_REP_UNR_CLO)
  
  # 2)
  df_sum = anova_summary_to_dataframe(mod)
  df_sum_REP = anova_summary_to_dataframe(mod_REP)
  df_sum_UNR = anova_summary_to_dataframe(mod_UNR)
  df_sum_CLO = anova_summary_to_dataframe(mod_CLO)
  df_sum_REP_UNR = anova_summary_to_dataframe(mod_REP_UNR)
  df_sum_REP_CLO = anova_summary_to_dataframe(mod_REP_CLO)
  df_sum_UNR_CLO = anova_summary_to_dataframe(mod_UNR_CLO)
  df_sum_REP_UNR_CLO = anova_summary_to_dataframe(mod_REP_UNR_CLO)
  
  Tuk = TukeyHSD(mod, conf.level=.95)
  Tuk_REP = TukeyHSD(mod_REP, conf.level=.95)
  Tuk_UNR = TukeyHSD(mod_UNR, conf.level=.95)
  Tuk_CLO = TukeyHSD(mod_CLO, conf.level=.95)
  Tuk_REP_UNR = TukeyHSD(mod_REP_UNR, conf.level=.95)
  Tuk_REP_CLO = TukeyHSD(mod_REP_CLO, conf.level=.95)
  Tuk_UNR_CLO = TukeyHSD(mod_UNR_CLO, conf.level=.95)
  Tuk_REP_UNR_CLO = TukeyHSD(mod_REP_UNR_CLO, conf.level=.95)
  
  # 3)
  df_Tuk = as.data.frame(Tuk$level)
  df_Tuk$comparison <- rownames(df_Tuk)
  df_Tuk <- df_Tuk[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk) <- NULL
  
  df_Tuk_REP = as.data.frame(Tuk_REP$level)
  df_Tuk_REP$comparison <- rownames(df_Tuk_REP)
  df_Tuk_REP <- df_Tuk_REP[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP) <- NULL
  
  df_Tuk_UNR = as.data.frame(Tuk_UNR$level)
  df_Tuk_UNR$comparison <- rownames(df_Tuk_UNR)
  df_Tuk_UNR <- df_Tuk_UNR[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_UNR) <- NULL
  
  df_Tuk_CLO = as.data.frame(Tuk_CLO$level)
  df_Tuk_CLO$comparison <- rownames(df_Tuk_CLO)
  df_Tuk_CLO <- df_Tuk_CLO[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_CLO) <- NULL
  
  df_Tuk_REP_UNR = as.data.frame(Tuk_REP_UNR$level)
  df_Tuk_REP_UNR$comparison <- rownames(df_Tuk_REP_UNR)
  df_Tuk_REP_UNR <- df_Tuk_REP_UNR[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_UNR) <- NULL
  
  df_Tuk_REP_CLO = as.data.frame(Tuk_REP_CLO$level)
  df_Tuk_REP_CLO$comparison <- rownames(df_Tuk_REP_CLO)
  df_Tuk_REP_CLO <- df_Tuk_REP_CLO[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_CLO) <- NULL
  
  df_Tuk_UNR_CLO = as.data.frame(Tuk_UNR_CLO$level)
  df_Tuk_UNR_CLO$comparison <- rownames(df_Tuk_UNR_CLO)
  df_Tuk_UNR_CLO <- df_Tuk_UNR_CLO[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_UNR_CLO) <- NULL
  
  df_Tuk_REP_UNR_CLO = as.data.frame(Tuk_REP_UNR_CLO$level)
  df_Tuk_REP_UNR_CLO$comparison <- rownames(df_Tuk_REP_UNR_CLO)
  df_Tuk_REP_UNR_CLO <- df_Tuk_REP_UNR_CLO[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_UNR_CLO) <- NULL
  
  # Create a list of dataframes
  df_list = list(df_cross_validation = df_cross_validation, 
                 df_sum = df_sum, 
                 df_Tuk = df_Tuk, 
                 df_sum_REP = df_sum_REP, 
                 df_Tuk_REP = df_Tuk_REP, 
                 df_sum_UNR = df_sum_UNR, 
                 df_Tuk_UNR = df_Tuk_UNR, 
                 df_sum_CLO = df_sum_CLO,                
                 df_Tuk_CLO = df_Tuk_CLO, 
                 df_sum_REP_UNR = df_sum_REP_UNR, 
                 df_Tuk_REP_UNR = df_Tuk_REP_UNR, 
                 df_sum_REP_CLO = df_sum_REP_CLO, 
                 df_Tuk_REP_CLO = df_Tuk_REP_CLO, 
                 df_sum_UNR_CLO = df_sum_UNR_CLO, 
                 df_Tuk_UNR_CLO = df_Tuk_UNR_CLO, 
                 df_sum_REP_UNR_CLO = df_sum_REP_UNR_CLO, 
                 df_Tuk_REP_UNR_CLO = df_Tuk_REP_UNR_CLO
  )
  
  # Export every df_2 to its own Excel sheet in one Excel file
  require(openxlsx)
  pathway_id_results_2 = paste(pathway_id, 'results_', id, '_2.xlsx', sep='')
  wb = createWorkbook()
  
  addWorksheet(wb, "cross_validation")
  writeData(wb, "cross_validation", df_cross_validation)
  
  ncol_agg = ncol(df_agg)
  ncol_sum = ncol(df_sum)
  addWorksheet(wb, "df")
  writeData(wb, "df", df_Tuk, startCol = 1, startRow = 1)
  writeData(wb, "df", df_sum, startCol = 7, startRow = 1)
  writeData(wb, "df", df_agg, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "REP")
  writeData(wb, "REP", df_Tuk_REP, startCol = 1, startRow = 1)
  writeData(wb, "REP", df_sum_REP, startCol = 7, startRow = 1)
  writeData(wb, "REP", df_agg_REP, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "UNR")
  writeData(wb, "UNR", df_Tuk_UNR, startCol = 1, startRow = 1)
  writeData(wb, "UNR", df_sum_UNR, startCol = 7, startRow = 1)
  writeData(wb, "UNR", df_agg_UNR, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "CLO")
  writeData(wb, "CLO", df_Tuk_CLO, startCol = 1, startRow = 1)
  writeData(wb, "CLO", df_sum_CLO, startCol = 7, startRow = 1)
  writeData(wb, "CLO", df_agg_CLO, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "REP_UNR")
  writeData(wb, "REP_UNR", df_Tuk_REP_UNR, startCol = 1, startRow = 1)
  writeData(wb, "REP_UNR", df_sum_REP_UNR, startCol = 7, startRow = 1)
  writeData(wb, "REP_UNR", df_agg_REP_UNR, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "REP_CLO")
  writeData(wb, "REP_CLO", df_Tuk_REP_CLO, startCol = 1, startRow = 1)
  writeData(wb, "REP_CLO", df_sum_REP_CLO, startCol = 7, startRow = 1)
  writeData(wb, "REP_CLO", df_agg_REP_CLO, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "UNR_CLO")
  writeData(wb, "UNR_CLO", df_Tuk_UNR_CLO, startCol = 1, startRow = 1)
  writeData(wb, "UNR_CLO", df_sum_UNR_CLO, startCol = 7, startRow = 1)
  writeData(wb, "UNR_CLO", df_agg_UNR_CLO, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "REP_UNR_CLO")
  writeData(wb, "REP_UNR_CLO", df_Tuk_REP_UNR_CLO, startCol = 1, startRow = 1)
  writeData(wb, "REP_UNR_CLO", df_sum_REP_UNR_CLO, startCol = 7, startRow = 1)
  writeData(wb, "REP_UNR_CLO", df_agg_REP_UNR_CLO, startCol = 14, startRow = 1)
  
  saveWorkbook(wb, pathway_id_results_2, overwrite = TRUE)  # Save the workbook to the specified pathway
  
}

export_results_3 = function(talk_id){
  library(readxl)
  library(xlsx)
  library(openxlsx)
  library(caret)
  
  pathway = '' # add pathway here
  id = as.character(talk_id)
  
  pathway_id = paste(pathway, id, '/', sep='')
  pathway_id_R = paste(pathway_id, 'R_', id, '.xlsx', sep='')
  
  df = read_excel(pathway_id_R)
  df = na.omit(df)
  df$related_class <- cut(df$related, breaks = seq(0, 5, by = 0.5), labels = c("0-0.5", "0.5-1", "1-1.5", "1.5-2", "2-2.5", "2.5-3", "3-3.5", "3.5-4", "4-4.5", "4.5-5"))
  
  # Complete model
  agg_ans = aggregate(alignment ~ answered, data = df, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  agg_rel = aggregate(alignment ~ related_class, data = df, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod = lm(alignment ~ answered + related, data = df)
  
  # Subset 1 variable
  df_REP = subset(df, repetition == 'no')
  agg_ans_REP = aggregate(alignment ~ answered, data = df_REP, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  agg_rel_REP = aggregate(alignment ~ related_class, data = df_REP, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP = lm(alignment ~ answered + related, data = df_REP)
  
  df_UNR = subset(df, relation_type_expected %in% c('Related', 'Hypophora'))
  agg_ans_UNR = aggregate(alignment ~ answered, data = df_UNR, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  agg_rel_UNR = aggregate(alignment ~ related_class, data = df_UNR, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_UNR = lm(alignment ~ answered + related, data = df_UNR)
  
  df_CLO = subset(df, question_type == 'open')
  agg_ans_CLO = aggregate(alignment ~ answered, data = df_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  agg_rel_CLO = aggregate(alignment ~ related_class, data = df_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_CLO = lm(alignment ~ answered + related, data = df_CLO)
  
  # Subset 2 variables
  df_REP_UNR = subset(df, repetition == 'no' & relation_type_expected %in% c('Related', 'Hypophora'))
  agg_ans_REP_UNR = aggregate(alignment ~ answered, data = df_REP_UNR, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  agg_rel_REP_UNR = aggregate(alignment ~ related_class, data = df_REP_UNR, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_UNR = lm(alignment ~ answered + related, data = df_REP_UNR)
  
  df_REP_CLO = subset(df, repetition == 'no' & question_type == 'open')
  agg_ans_REP_CLO = aggregate(alignment ~ answered, data = df_REP_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  agg_rel_REP_CLO = aggregate(alignment ~ related_class, data = df_REP_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_CLO = lm(alignment ~ answered + related, data = df_REP_CLO)
  
  df_UNR_CLO = subset(df, relation_type_expected %in% c('Related', 'Hypophora') & question_type == 'open')
  agg_ans_UNR_CLO = aggregate(alignment ~ answered, data = df_UNR_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  agg_rel_UNR_CLO = aggregate(alignment ~ related_class, data = df_UNR_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_UNR_CLO = lm(alignment ~ answered + related, data = df_UNR_CLO)
  
  # Subset all 3 variables
  df_REP_UNR_CLO = subset(df, repetition == 'no' & relation_type_expected %in% c('Related', 'Hypophora') & question_type == 'open')
  agg_ans_REP_UNR_CLO = aggregate(alignment ~ answered, data = df_REP_UNR_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  agg_rel_REP_UNR_CLO = aggregate(alignment ~ related_class, data = df_REP_UNR_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_UNR_CLO = lm(alignment ~ answered + related, data = df_REP_UNR_CLO)
  
  # Cross-validation for finding the best model
  cross_validate = train(alignment ~ answered + related, data = df, method = "lm")
  cross_validate_REP = train(alignment ~ answered + related, data = df_REP, method = "lm")
  cross_validate_UNR = train(alignment ~ answered + related, data = df_UNR, method = "lm")
  cross_validate_CLO = train(alignment ~ answered + related, data = df_CLO, method = "lm")
  cross_validate_REP_UNR = train(alignment ~ answered + related, data = df_REP_UNR, method = "lm")
  cross_validate_REP_CLO = train(alignment ~ answered + related, data = df_REP_CLO, method = "lm")
  cross_validate_UNR_CLO = train(alignment ~ answered + related, data = df_UNR_CLO, method = "lm")
  cross_validate_REP_UNR_CLO = train(alignment ~ answered + related, data = df_REP_UNR_CLO, method = "lm")
  
  #view summary of k-fold CV   
  df_CV = as.data.frame(cross_validate$results)
  df_CV_REP = as.data.frame = (cross_validate_REP$results)
  df_CV_UNR = as.data.frame = (cross_validate_UNR$results)
  df_CV_CLO = as.data.frame = (cross_validate_CLO$results)
  df_CV_REP_UNR = as.data.frame = (cross_validate_REP_UNR$results)
  df_CV_REP_CLO = as.data.frame = (cross_validate_REP_CLO$results)
  df_CV_UNR_CLO = as.data.frame = (cross_validate_UNR_CLO$results)
  df_CV_REP_UNR_CLO = as.data.frame = (cross_validate_REP_UNR_CLO$results)
  
  df_cross_validation_1 = bind_rows(df_CV, 
                                    df_CV_REP, 
                                    df_CV_UNR, 
                                    df_CV_CLO, 
                                    df_CV_REP_UNR, 
                                    df_CV_REP_CLO, 
                                    df_CV_UNR_CLO, 
                                    df_CV_REP_UNR_CLO)
  
  names = c("df_CV", "df_CV_REP", "df_CV_UNR", "df_CV_CLO", "df_CV_REP_UNR", "df_CV_REP_CLO", "df_CV_UNR_CLO", "df_CV_REP_UNR_CLO")
  df_names = as.data.frame(names)
  df_cross_validation = bind_cols(df_names, df_cross_validation_1)
  
  # Create dataframes of 1) the aggregate summary stats and 2) lm summaries
  # 1) ans
  df_agg_ans = as.data.frame(agg_ans)
  df_agg_ans_REP = as.data.frame(agg_ans_REP)
  df_agg_ans_UNR = as.data.frame(agg_ans_UNR)
  df_agg_ans_CLO = as.data.frame(agg_ans_CLO)
  df_agg_ans_REP_UNR = as.data.frame(agg_ans_REP_UNR)
  df_agg_ans_REP_CLO = as.data.frame(agg_ans_REP_CLO)
  df_agg_ans_UNR_CLO = as.data.frame(agg_ans_UNR_CLO)
  df_agg_ans_REP_UNR_CLO = as.data.frame(agg_ans_REP_UNR_CLO)
  
  # 1) rel
  df_agg_rel = as.data.frame(agg_rel)
  df_agg_rel_REP = as.data.frame(agg_rel_REP)
  df_agg_rel_UNR = as.data.frame(agg_rel_UNR)
  df_agg_rel_CLO = as.data.frame(agg_rel_CLO)
  df_agg_rel_REP_UNR = as.data.frame(agg_rel_REP_UNR)
  df_agg_rel_REP_CLO = as.data.frame(agg_rel_REP_CLO)
  df_agg_rel_UNR_CLO = as.data.frame(agg_rel_UNR_CLO)
  df_agg_rel_REP_UNR_CLO = as.data.frame(agg_rel_REP_UNR_CLO)
  
  # 2)
  df_sum = lm_summary_to_dataframe(mod)
  df_sum_REP = lm_summary_to_dataframe(mod_REP)
  df_sum_UNR = lm_summary_to_dataframe(mod_UNR)
  df_sum_CLO = lm_summary_to_dataframe(mod_CLO)
  df_sum_REP_UNR = lm_summary_to_dataframe(mod_REP_UNR)
  df_sum_REP_CLO = lm_summary_to_dataframe(mod_REP_CLO)
  df_sum_UNR_CLO = lm_summary_to_dataframe(mod_UNR_CLO)
  df_sum_REP_UNR_CLO = lm_summary_to_dataframe(mod_REP_UNR_CLO)
  
  # Create a list of dataframes
  df_list = list(df_cross_validation = df_cross_validation, 
                 df_sum = df_sum, 
                 df_sum_REP = df_sum_REP, 
                 df_sum_UNR = df_sum_UNR, 
                 df_sum_CLO = df_sum_CLO, 
                 df_sum_REP_UNR = df_sum_REP_UNR, 
                 df_sum_REP_CLO = df_sum_REP_CLO, 
                 df_sum_UNR_CLO = df_sum_UNR_CLO, 
                 df_sum_REP_UNR_CLO = df_sum_REP_UNR_CLO
  )
  
  # Export every df_3 to its own Excel sheet in one Excel file
  require(openxlsx)
  pathway_id_results_3 = paste(pathway_id, 'results_', id, '_3.xlsx', sep='')
  wb = createWorkbook()
  
  addWorksheet(wb, "cross_validation")
  writeData(wb, "cross_validation", df_cross_validation)
  
  ncol_agg_ans = ncol(df_agg_ans)
  ncol_agg_rel = ncol(df_agg_rel)
  ncol_sum = ncol(df_sum)
  addWorksheet(wb, "df")
  writeData(wb, "df", df_sum, startCol = 1, startRow = 1)
  writeData(wb, "df", df_agg_ans, startCol = 7, startRow = 1)
  writeData(wb, "df", df_agg_rel, startCol = 10, startRow = 1)
  
  addWorksheet(wb, "REP")
  writeData(wb, "REP", df_sum_REP, startCol = 1, startRow = 1)
  writeData(wb, "REP", df_agg_ans_REP, startCol = 7, startRow = 1)
  writeData(wb, "REP", df_agg_rel_REP, startCol = 10, startRow = 1)
  
  addWorksheet(wb, "UNR")
  writeData(wb, "UNR", df_sum_UNR, startCol = 1, startRow = 1)
  writeData(wb, "UNR", df_agg_ans_UNR, startCol = 7, startRow = 1)
  writeData(wb, "UNR", df_agg_rel_UNR, startCol = 10, startRow = 1)
  
  addWorksheet(wb, "CLO")
  writeData(wb, "CLO", df_sum_CLO, startCol = 1, startRow = 1)
  writeData(wb, "CLO", df_agg_ans_CLO, startCol = 7, startRow = 1)
  writeData(wb, "CLO", df_agg_rel_CLO, startCol = 10, startRow = 1)
  
  addWorksheet(wb, "REP_UNR")
  writeData(wb, "REP_UNR", df_sum_REP_UNR, startCol = 1, startRow = 1)
  writeData(wb, "REP_UNR", df_agg_ans_REP_UNR, startCol = 7, startRow = 1)
  writeData(wb, "REP_UNR", df_agg_rel_REP_UNR, startCol = 10, startRow = 1)
  
  addWorksheet(wb, "REP_CLO")
  writeData(wb, "REP_CLO", df_sum_REP_CLO, startCol = 1, startRow = 1)
  writeData(wb, "REP_CLO", df_agg_ans_REP_CLO, startCol = 7, startRow = 1)
  writeData(wb, "REP_CLO", df_agg_rel_REP_CLO, startCol = 10, startRow = 1)
  
  addWorksheet(wb, "UNR_CLO")
  writeData(wb, "UNR_CLO", df_sum_UNR_CLO, startCol = 1, startRow = 1)
  writeData(wb, "UNR_CLO", df_agg_ans_UNR_CLO, startCol = 7, startRow = 1)
  writeData(wb, "UNR_CLO", df_agg_rel_UNR_CLO, startCol = 10, startRow = 1)
  
  addWorksheet(wb, "REP_UNR_CLO")
  writeData(wb, "REP_UNR_CLO", df_sum_REP_UNR_CLO, startCol = 1, startRow = 1)
  writeData(wb, "REP_UNR_CLO", df_agg_ans_REP_UNR_CLO, startCol = 7, startRow = 1)
  writeData(wb, "REP_UNR_CLO", df_agg_rel_REP_UNR_CLO, startCol = 10, startRow = 1)
  
  saveWorkbook(wb, pathway_id_results_3, overwrite = TRUE)  # Save the workbook to the specified pathway
  
}
  
export_results_4 = function(talk_id){
  library(readxl)
  library(xlsx)
  library(openxlsx)
  library(caret)
  
  pathway = '' # add pathway here
  id = as.character(talk_id)
  
  pathway_id = paste(pathway, id, '/', sep='')
  pathway_id_R = paste(pathway_id, 'R_', id, '.xlsx', sep='')
  
  df = read_excel(pathway_id_R)
  df = na.omit(df)
  
  # Complete model
  agg = aggregate(alignment ~ relation_type, data = df, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod <- aov(alignment ~ relation_type, data=df)
  
  # Subset 1 variable
  df_REP = subset(df, repetition == 'no')
  agg_REP = aggregate(alignment ~ relation_type, data = df_REP, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP = aov(alignment ~ relation_type, data=df_REP)
  
  df_UNR = subset(df, relation_type_expected %in% c('Related', 'Hypophora'))
  agg_UNR = aggregate(alignment ~ relation_type, data = df_UNR, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_UNR = aov(alignment ~ relation_type, data=df_UNR)
  
  df_CLO = subset(df, question_type == 'open')
  agg_CLO = aggregate(alignment ~ relation_type, data = df_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_CLO = aov(alignment ~ relation_type, data=df_CLO)
  
  # Subset 2 variables
  df_REP_UNR = subset(df, repetition == 'no' & relation_type_expected %in% c('Related', 'Hypophora'))
  agg_REP_UNR = aggregate(alignment ~ relation_type, data = df_REP_UNR, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_UNR = aov(alignment ~ relation_type, data=df_REP_UNR)
  
  df_REP_CLO = subset(df, repetition == 'no' & question_type == 'open')
  agg_REP_CLO = aggregate(alignment ~ relation_type, data = df_REP_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_CLO = aov(alignment ~ relation_type, data=df_REP_CLO)
  
  df_UNR_CLO = subset(df, relation_type_expected %in% c('Related', 'Hypophora') & question_type == 'open')
  agg_UNR_CLO = aggregate(alignment ~ relation_type, data = df_UNR_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_UNR_CLO = aov(alignment ~ relation_type, data=df_UNR_CLO)
  
  # Subset all 3 variables
  df_REP_UNR_CLO = subset(df, repetition == 'no' & relation_type_expected %in% c('Related', 'Hypophora') & question_type == 'open')
  agg_REP_UNR_CLO = aggregate(alignment ~ relation_type, data = df_REP_UNR_CLO, FUN = function(x) c(mean = mean(x), sd = sd(x)))
  mod_REP_UNR_CLO = aov(alignment ~ relation_type, data=df_REP_UNR_CLO)
  
  # Cross-validation for finding the best model
  cross_validate = train(alignment ~ relation_type, data = df, method = "lm")
  cross_validate_REP = train(alignment ~ relation_type, data = df_REP, method = "lm")
  cross_validate_UNR = train(alignment ~ relation_type, data = df_UNR, method = "lm")
  cross_validate_CLO = train(alignment ~ relation_type, data = df_CLO, method = "lm")
  cross_validate_REP_UNR = train(alignment ~ relation_type, data = df_REP_UNR, method = "lm")
  cross_validate_REP_CLO = train(alignment ~ relation_type, data = df_REP_CLO, method = "lm")
  cross_validate_UNR_CLO = train(alignment ~ relation_type, data = df_UNR_CLO, method = "lm")
  cross_validate_REP_UNR_CLO = train(alignment ~ relation_type, data = df_REP_UNR_CLO, method = "lm")
  
  #view summary of k-fold CV   
  df_CV = as.data.frame(cross_validate$results)
  df_CV_REP = as.data.frame = (cross_validate_REP$results)
  df_CV_UNR = as.data.frame = (cross_validate_UNR$results)
  df_CV_CLO = as.data.frame = (cross_validate_CLO$results)
  df_CV_REP_UNR = as.data.frame = (cross_validate_REP_UNR$results)
  df_CV_REP_CLO = as.data.frame = (cross_validate_REP_CLO$results)
  df_CV_UNR_CLO = as.data.frame = (cross_validate_UNR_CLO$results)
  df_CV_REP_UNR_CLO = as.data.frame = (cross_validate_REP_UNR_CLO$results)
  
  df_cross_validation_1 = bind_rows(df_CV, 
                                    df_CV_REP, 
                                    df_CV_UNR, 
                                    df_CV_CLO, 
                                    df_CV_REP_UNR, 
                                    df_CV_REP_CLO, 
                                    df_CV_UNR_CLO, 
                                    df_CV_REP_UNR_CLO)
  
  names = c("df_CV", "df_CV_REP", "df_CV_UNR", "df_CV_CLO", "df_CV_REP_UNR", "df_CV_REP_CLO", "df_CV_UNR_CLO", "df_CV_REP_UNR_CLO")
  df_names = as.data.frame(names)
  df_cross_validation = bind_cols(df_names, df_cross_validation_1)
  
  # Create dataframes of 1) the aggregate summary stats, 2) anova summaries and 3) TukeyHSDs
  # 1)
  df_agg = as.data.frame(agg)
  df_agg_REP = as.data.frame(agg_REP)
  df_agg_UNR = as.data.frame(agg_UNR)
  df_agg_CLO = as.data.frame(agg_CLO)
  df_agg_REP_UNR = as.data.frame(agg_REP_UNR)
  df_agg_REP_CLO = as.data.frame(agg_REP_CLO)
  df_agg_UNR_CLO = as.data.frame(agg_UNR_CLO)
  df_agg_REP_UNR_CLO = as.data.frame(agg_REP_UNR_CLO)
  
  df_sum = anova_summary_to_dataframe(mod)
  df_sum_REP = anova_summary_to_dataframe(mod_REP)
  df_sum_UNR = anova_summary_to_dataframe(mod_UNR)
  df_sum_CLO = anova_summary_to_dataframe(mod_CLO)
  df_sum_REP_UNR = anova_summary_to_dataframe(mod_REP_UNR)
  df_sum_REP_CLO = anova_summary_to_dataframe(mod_REP_CLO)
  df_sum_UNR_CLO = anova_summary_to_dataframe(mod_UNR_CLO)
  df_sum_REP_UNR_CLO = anova_summary_to_dataframe(mod_REP_UNR_CLO)
  
  Tuk = TukeyHSD(mod, conf.level=.95)
  Tuk_REP = TukeyHSD(mod_REP, conf.level=.95)
  Tuk_UNR = TukeyHSD(mod_UNR, conf.level=.95)
  Tuk_CLO = TukeyHSD(mod_CLO, conf.level=.95)
  Tuk_REP_UNR = TukeyHSD(mod_REP_UNR, conf.level=.95)
  Tuk_REP_CLO = TukeyHSD(mod_REP_CLO, conf.level=.95)
  Tuk_UNR_CLO = TukeyHSD(mod_UNR_CLO, conf.level=.95)
  Tuk_REP_UNR_CLO = TukeyHSD(mod_REP_UNR_CLO, conf.level=.95)
  
  df_Tuk = as.data.frame(Tuk$relation_type)
  df_Tuk$comparison <- rownames(df_Tuk)
  df_Tuk <- df_Tuk[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk) <- NULL
  
  df_Tuk_REP = as.data.frame(Tuk_REP$relation_type)
  df_Tuk_REP$comparison <- rownames(df_Tuk_REP)
  df_Tuk_REP <- df_Tuk_REP[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP) <- NULL
  
  df_Tuk_UNR = as.data.frame(Tuk_UNR$relation_type)
  df_Tuk_UNR$comparison <- rownames(df_Tuk_UNR)
  df_Tuk_UNR <- df_Tuk_UNR[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_UNR) <- NULL
  
  df_Tuk_CLO = as.data.frame(Tuk_CLO$relation_type)
  df_Tuk_CLO$comparison <- rownames(df_Tuk_CLO)
  df_Tuk_CLO <- df_Tuk_CLO[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_CLO) <- NULL
  
  df_Tuk_REP_UNR = as.data.frame(Tuk_REP_UNR$relation_type)
  df_Tuk_REP_UNR$comparison <- rownames(df_Tuk_REP_UNR)
  df_Tuk_REP_UNR <- df_Tuk_REP_UNR[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_UNR) <- NULL
  
  df_Tuk_REP_CLO = as.data.frame(Tuk_REP_CLO$relation_type)
  df_Tuk_REP_CLO$comparison <- rownames(df_Tuk_REP_CLO)
  df_Tuk_REP_CLO <- df_Tuk_REP_CLO[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_CLO) <- NULL
  
  df_Tuk_UNR_CLO = as.data.frame(Tuk_UNR_CLO$relation_type)
  df_Tuk_UNR_CLO$comparison <- rownames(df_Tuk_UNR_CLO)
  df_Tuk_UNR_CLO <- df_Tuk_UNR_CLO[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_UNR_CLO) <- NULL
  
  df_Tuk_REP_UNR_CLO = as.data.frame(Tuk_REP_UNR_CLO$relation_type)
  df_Tuk_REP_UNR_CLO$comparison <- rownames(df_Tuk_REP_UNR_CLO)
  df_Tuk_REP_UNR_CLO <- df_Tuk_REP_UNR_CLO[, c("comparison", "diff", "lwr", "upr", "p adj")]
  rownames(df_Tuk_REP_UNR_CLO) <- NULL
  
  # Create a list of dataframes
  df_list = list(df_cross_validation = df_cross_validation, 
                 df_sum = df_sum, 
                 df_Tuk = df_Tuk, 
                 df_sum_REP = df_sum_REP, 
                 df_Tuk_REP = df_Tuk_REP, 
                 df_sum_UNR = df_sum_UNR, 
                 df_Tuk_UNR = df_Tuk_UNR, 
                 df_sum_CLO = df_sum_CLO,                
                 df_Tuk_CLO = df_Tuk_CLO, 
                 df_sum_REP_UNR = df_sum_REP_UNR, 
                 df_Tuk_REP_UNR = df_Tuk_REP_UNR, 
                 df_sum_REP_CLO = df_sum_REP_CLO, 
                 df_Tuk_REP_CLO = df_Tuk_REP_CLO, 
                 df_sum_UNR_CLO = df_sum_UNR_CLO, 
                 df_Tuk_UNR_CLO = df_Tuk_UNR_CLO, 
                 df_sum_REP_UNR_CLO = df_sum_REP_UNR_CLO, 
                 df_Tuk_REP_UNR_CLO = df_Tuk_REP_UNR_CLO
  )
  
  # Export every df_4 to its own Excel sheet in one Excel file
  require(openxlsx)
  pathway_id_results_4 = paste(pathway_id, 'results_', id, '_4.xlsx', sep='')
  wb = createWorkbook()
  
  addWorksheet(wb, "cross_validation")
  writeData(wb, "cross_validation", df_cross_validation)
  
  ncol_agg = ncol(df_agg)
  ncol_sum = ncol(df_sum)
  addWorksheet(wb, "df")
  writeData(wb, "df", df_Tuk, startCol = 1, startRow = 1)
  writeData(wb, "df", df_sum, startCol = 7, startRow = 1)
  writeData(wb, "df", df_agg, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "REP")
  writeData(wb, "REP", df_Tuk_REP, startCol = 1, startRow = 1)
  writeData(wb, "REP", df_sum_REP, startCol = 7, startRow = 1)
  writeData(wb, "REP", df_agg_REP, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "UNR")
  writeData(wb, "UNR", df_Tuk_UNR, startCol = 1, startRow = 1)
  writeData(wb, "UNR", df_sum_UNR, startCol = 7, startRow = 1)
  writeData(wb, "UNR", df_agg_UNR, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "CLO")
  writeData(wb, "CLO", df_Tuk_CLO, startCol = 1, startRow = 1)
  writeData(wb, "CLO", df_sum_CLO, startCol = 7, startRow = 1)
  writeData(wb, "CLO", df_agg_CLO, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "REP_UNR")
  writeData(wb, "REP_UNR", df_Tuk_REP_UNR, startCol = 1, startRow = 1)
  writeData(wb, "REP_UNR", df_sum_REP_UNR, startCol = 7, startRow = 1)
  writeData(wb, "REP_UNR", df_agg_REP_UNR, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "REP_CLO")
  writeData(wb, "REP_CLO", df_Tuk_REP_CLO, startCol = 1, startRow = 1)
  writeData(wb, "REP_CLO", df_sum_REP_CLO, startCol = 7, startRow = 1)
  writeData(wb, "REP_CLO", df_agg_REP_CLO, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "UNR_CLO")
  writeData(wb, "UNR_CLO", df_Tuk_UNR_CLO, startCol = 1, startRow = 1)
  writeData(wb, "UNR_CLO", df_sum_UNR_CLO, startCol = 7, startRow = 1)
  writeData(wb, "UNR_CLO", df_agg_UNR_CLO, startCol = 14, startRow = 1)
  
  addWorksheet(wb, "REP_UNR_CLO")
  writeData(wb, "REP_UNR_CLO", df_Tuk_REP_UNR_CLO, startCol = 1, startRow = 1)
  writeData(wb, "REP_UNR_CLO", df_sum_REP_UNR_CLO, startCol = 7, startRow = 1)
  writeData(wb, "REP_UNR_CLO", df_agg_REP_UNR_CLO, startCol = 14, startRow = 1)
  
  saveWorkbook(wb, pathway_id_results_4, overwrite = TRUE)  # Save the workbook to the specified pathway
  
}

export_results_1(1927)
export_results_2(1927)
export_results_3(1927)
export_results_4(1927)

export_results_1(1971)
export_results_2(1971)
export_results_3(1971)
export_results_4(1971)

export_results_1(1976)
export_results_2(1976)
export_results_3(1976)
export_results_4(1976)

export_results_1("Total")
export_results_2("Total")
export_results_3("Total")
export_results_4("Total")
