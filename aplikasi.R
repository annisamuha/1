# Step 1: Read Data and Create Variables
read_and_create_variables <- function(file_path) {
  library(readxl)
  data <- read_excel(file_path)
  data$Date <- as.factor(data$Date)
  for (i in 1:30) {
    assign(paste0("Y", i), data[[paste0("Y", i)]])
  }
  
  X <- data$X
  BI_RATE <- data$BI_Rate_bulanan
  Rbr <- mean(BI_RATE)
  
  return(list(data = data, X = X, BI_RATE = BI_RATE, Rbr = Rbr,Y1=Y1,Y2=Y2,Y3=Y3,Y4=Y4,
              Y5=Y5,Y6=Y6,Y7=Y7,Y8=Y8,Y9=Y9,Y10=Y10,Y11=Y11,Y12=Y12,Y13=Y13,Y14=Y14,Y15=Y15,Y16=Y16,Y17=Y17,Y18=Y18,Y19=Y19,Y20=Y20,Y21=Y21,Y22=Y22,Y23=Y23,Y24=Y24,Y25=Y25,Y26=Y26,Y27=Y27,Y28=Y28,Y29=Y29,Y30=Y30))
}
# Step 2: Calculate E(Ri) and E(RM)
calculate_means <- function(data, Rbr) {
  variabel_tersisa <- c()
  mean_values <- c()
  
  for (i in 1:30) {
    mean_value <- mean(data[[paste0("Y", i)]])
    if (mean_value > Rbr) {
      cat(paste0("Mean Y", i, ": ", mean_value, " lebih besar daripada Rbr: ", Rbr, "\n"))
      
      variabel_tersisa <- c(variabel_tersisa, paste0("Y", i))
    } else {
      cat(paste0("Mean Y", i, ": ", mean_value, " tidak lebih besar daripada Rbr ", Rbr, "\n"))
      # Ganti dengan nilai yang sesuai, meskipun tidak relevan
    }
    mean_values <- c(mean_values, mean_value)
  }
  
  # Pastikan panjang kedua vektor sama sebelum membuat data.frame
  result_df <- data.frame(Variabel = paste0("Y", 1:30), Mean_Value = mean_values)
  return(list(result_df = result_df, variabel_tersisa = variabel_tersisa))
}

# Step 3: Calculate Residuals
calculate_residuals <- function(variabel_tersisa, data) {
  nilai_fit_list <- list()
  residuals_list <- list()
  
  for (i in variabel_tersisa) {
    nilai_fit <- lm(paste(i, "~ X"), data = data)
    nilai_fit_list[[i]] <- nilai_fit
    residuals_list[[paste0("e", i)]] <- residuals(nilai_fit)
  }
  
  return(list(nilai_fit_list = nilai_fit_list, residuals_list = residuals_list))
}

# Step 4: Test Covariance Assumption
test_covariance_assumption <- function(residuals_list) {
  matriksdata <- do.call(cbind, residuals_list)
  mcovdata <- cov(matriksdata)
  mcovdata <- round(mcovdata)
  return(data.frame(Covarian = rownames(mcovdata), mcovdata))
}

# Step 5: Test Covariance Assumption with IHSG
test_covariance_with_ihsg <- function(variabel_tersisa, residuals_list, X) {
  kovarian_list <- list()
  
  for (i in variabel_tersisa) {
    kovarian_res_X <- round(cov(residuals_list[[paste0("e", i)]], X, method = "pearson"))
    kovarian_list[[paste0("kov_", i, "_X")]] <- kovarian_res_X
    
    cat(paste0("Kov_e", i, "_X:\n"))
    print(kovarian_res_X)
  }
  df_kovarian <- data.frame(
    Variabel = names(kovarian_list),
    Nilai = unlist(kovarian_list)
  )
  return(df_kovarian)
}

# Step 6: Calculate Alpha and Beta
calculate_alpha_beta <- function(variabel_tersisa, data) {
  alpha_beta_list <- list()
  alpha_list <- list()
  beta_list <- list()
  for (i in variabel_tersisa) {
    if (i %in% colnames(data)) {
      X_Y <- matrix(c(data$X, data[[i]]), ncol = 2)
      mcov <- var(X_Y)
      cov_X_Y <- mcov[2, 1]
      var_Y <- mcov[2, 2]
      var_X <- mcov[1, 1]
      beta_Y <- cov_X_Y / var_X
      alpha_Y <- mean(data[[i]]) - (beta_Y * mean(data$X))
      
      cat(paste0("beta_", i, "=", beta_Y, "\n"))
      cat(paste0("alpha_", i, "=", alpha_Y, "\n"))
      
      alpha_list[[paste0("alpha_", i)]] <- alpha_Y
      beta_list[[paste0("beta_", i)]] <- beta_Y
      alpha_beta_list[[paste0("alpha_", i)]] <- alpha_Y
      alpha_beta_list[[paste0("beta_", i)]] <- beta_Y
    } else {
      warning(paste("Column", i, "not found in the dataset"))
    }
  }
  df_alpha_beta<- data.frame(
    Variabel = variabel_tersisa,
    Alpha = unlist(alpha_list),
    Beta = unlist(beta_list)
  )
  return(list(df_alpha_beta = df_alpha_beta, alpha_beta_list = alpha_beta_list,alpha_list = alpha_list, beta_list = beta_list))
}

# Step 7: Calculate ERB
calculate_erb <- function(variabel_tersisa, alpha_beta_list, data) {
  erb_list <- list()
  
  for (i in variabel_tersisa) {
    erb_Y <- (mean(data[[i]]) - step_1$Rbr) / alpha_beta_list[[paste0("beta_", i)]]
    
    cat(paste0("ERB_", i, "=", erb_Y, "\n"))
    
    erb_list[[paste0("ERB_", i)]] <- erb_Y
  }
  
  return(erb_list)
}

# Step 8: Sort ERB
sort_erb <- function(erb_list,variabel_tersisa) {
  nilai_erb_urut <- sort(unlist(erb_list), decreasing = TRUE)
  sorted_indices <- order(unlist(erb_list), decreasing = TRUE)
  data_ERB <- unlist(erb_list)[sorted_indices]
  variabel_tersisa <- variabel_tersisa[sorted_indices]
  return(list(nilai_erb_urut = nilai_erb_urut, data_ERB = data_ERB, variabel_tersisa = variabel_tersisa))
}

# Step 9: Calculate Ai and Bi
calculate_ai_bi <- function(variabel_tersisa, alpha_beta_list, residuals_list, data, Rbr) {
  ai_list <- list()
  bi_list <- list()
  
  for (i in variabel_tersisa) {
    A_Y <- ((mean(data[[i]]) - Rbr) * alpha_beta_list[[paste0("beta_", i)]]) / var(residuals_list[[paste0("e", i)]])
    B_Y <- ((alpha_beta_list[[paste0("beta_", i)]])^2) / var(residuals_list[[paste0("e", i)]])
    
    cat(paste0("A_", i, "=", A_Y, "\n"))
    cat(paste0("B_", i, "=", B_Y, "\n"))
    
    ai_list[[paste0("A_", i)]] <- A_Y
    bi_list[[paste0("B_", i)]] <- B_Y
  }
  
  return(list(ai_list = ai_list, bi_list = bi_list))
}

# Step 10: Calculate Ci
calculate_ci <- function(variabel_tersisa, ai_list,bi_list, X) {
  ci_list <- list()
  sum_ai=0
  sum_bi=0
  
  for (i in variabel_tersisa) {
    ai=ai_list[[paste0("A_", i)]]
    bi=bi_list[[paste0("B_", i)]]
    sum_ai=sum_ai+ai
    sum_bi=sum_bi+bi
    Ci_Y <- ((var(X)) * sum_ai / (1 + var(X) *sum_bi ))
    
    cat(paste0("C_", i, "=", Ci_Y, "\n"))
    
    ci_list[[paste0("C_", i)]] <- Ci_Y
  }
  
  return(ci_list)
}

# Step 11: Sort Variabel Tersisa
sort_variabel_tersisa <- function(variabel_tersisa, erb_list,ai_list,bi_list, ci_list) {
  #sorted_indices <- order(unlist(erb_list), decreasing = TRUE)
  data_ERB <- unlist(erb_list)
  data_Ci <- unlist(ci_list)
  data_ai <- unlist(ai_list)
  data_bi <- unlist(bi_list)
  variabel_tersisa <- variabel_tersisa
  
  cat("Variabel_tersisa yang sudah diurutkan:\n")
  print(variabel_tersisa)
  
  perbandingan <- data.frame(variable = variabel_tersisa, ERB = data_ERB,Ai=data_ai,Bi=data_bi, Ci = data_Ci)
  print(perbandingan)
}

# Step 12: Determine Variables in Portfolio
determine_variables_in_portfolio <- function(variabel_tersisa, ci_list, erb_list) {
  variabel_max_ci <- names(ci_list)[which.max(unlist(ci_list))]
  erb_df <- data.frame(variable = names(erb_list), ERB = unlist(erb_list))
  ERB_dipilih <- NULL
  
  for (i in seq_along(erb_df$variable)) {
    if (grepl(sub("ERB_", "", erb_df$variable[i]), variabel_max_ci)) {
      ERB_dipilih <- erb_df$ERB[i]
      break
    }
  }
  
  cat("Variabel Ci tertinggi: ", variabel_max_ci, "\n")
  cat("ERB: ", ERB_dipilih, "\n")
  
  variabel_elim <- erb_df$variable[erb_df$ERB < ERB_dipilih]
  variabel_elim <- gsub("ERB_", "", variabel_elim)
  variabel_tersisa <- setdiff(variabel_tersisa, variabel_elim)
  
  cat("Variabel yang tersisa: ", paste(variabel_tersisa, collapse = ", "), "\n")
  return(list(variabel_max_ci = variabel_max_ci, variabel_tersisa = variabel_tersisa))
}

determine_cut_off_point <- function(ci_list, variabel_max_ci) {
  C_bintang <- ci_list[[variabel_max_ci]]
  return(C_bintang)
}

# Step 13: Calculate Z
calculate_z <- function(variabel_tersisa, alpha_beta_list, erb_list, ci_list, C_bintang, residuals_list) {
  z_list <- list()
  sum_Z <- 0
  
  for (i in variabel_tersisa) {
    Z_Y <- (alpha_beta_list[[paste0("beta_", i)]] / var(residuals_list[[paste0("e", i)]]) *
              (erb_list[[paste0("ERB_", i)]] - C_bintang))
    
    cat(paste0("Z_", i, "=", Z_Y, "\n"))
    
    z_list[[paste0("Z_", i)]] <- Z_Y
    sum_Z <- sum_Z + Z_Y
  }
  
  cat("sum_Z=", sum_Z, "\n")
  return(z_list)
}

# Step 14: Calculate Weights
calculate_weights <- function(variabel_tersisa, z_list) {
  weights <- list()
  sum_Z <- sum(unlist(z_list))
  
  for (i in variabel_tersisa) {
    current_Z <- z_list[[paste0("Z_", i)]]
    weight_i <- current_Z / sum_Z
    
    cat(paste0("w_", i, "=", weight_i, "\n"))
    
    weights[[paste0("w_", i)]] <- weight_i
  }
  
  weights_df <- data.frame(variable = names(weights), weight = unlist(weights))
  return(weights_df)
}

# Step 15: Multivariate Normality Test

calculate_mahalanobis_test <- function(variabel_tersisa, data) {
  A <- matrix(ncol = length(variabel_tersisa), nrow = nrow(data))
  
  # Populate matrix A
  for (i in seq_along(variabel_tersisa)) {
    var_name <- variabel_tersisa[i]
    var_data <- data[[var_name]]
    
    if (length(var_data) != nrow(data)) {
      stop(paste("Panjang variabel", var_name, "tidak sesuai dengan panjang data"))
    }
    
    A[, i] <- var_data
  }
  
  dimnames(A) <- list(NULL, variabel_tersisa)
  
  # Calculate Mahalanobis distances
  rerata <- apply(A, 2, mean)
  mcov <- var(A)
  dj_kuadrat <- mahalanobis(A, center = rerata, cov = mcov)
  dj_kuadrat <- sort(dj_kuadrat)
  
  # ECDF calculations
  ecdf_empiris <- ecdf(dj_kuadrat)
  cdf_ds <- ecdf_empiris(dj_kuadrat)
  cdf_chisq <- pchisq(dj_kuadrat, df = ncol(A))
  
  n <- length(dj_kuadrat)
  D_bintang <- 1.36 / sqrt(n)
  
  # Calculate test statistic
  s_x <- cdf_ds
  f_x <- cdf_chisq
  diff_s_f <- abs(s_x - f_x)
  diff_s_minus_1_f <- c(NA, abs(s_x[-length(s_x)] - f_x[-1]))
  max_values <- pmax(diff_s_f, diff_s_minus_1_f, na.rm = TRUE)
  statistikuji <- max(max_values, na.rm = TRUE)
  
  # Display test result
  if (statistikuji < D_bintang) {
    cat("H0 diterima, Jarak Mahalanobis Berdistribusi Chi Kuadrat\n")
  } else {
    cat("H0 ditolak, Jarak Mahalanobis Tidak Berdistribusi Chi Kuadrat\n")
  }
  
  # Plot ECDF
  par(mfrow = c(1, 1))
  plot(cdf_ds, main = "Cumulative Distribution Function", ylim = c(0, 1))
  lines(pchisq(dj_kuadrat, df = ncol(A)), col = "purple")
  
  conclusion <- ifelse(statistikuji < D_bintang, "H0 diterima, Jarak Mahalanobis Berdistribusi Chi Kuadrat", "H0 ditolak, Jarak Mahalanobis Tidak Berdistribusi Chi Kuadrat")
  
  return(list(statistikuji = statistikuji, conclusion = conclusion, cdf_ds=cdf_ds, df=ncol(A),dj_kuadrat=dj_kuadrat,D_bintang=D_bintang))
}


# Step 16: Calculate VaR (1 Simulation)
calculate_var_1_simulation <- function(data, variabel_tersisa, weights_df) {
  library("MASS")
  A <- matrix(nrow = nrow(data), ncol = length(variabel_tersisa))
  
  for (i in seq_along(variabel_tersisa)) {
    A[, i] <- data[[variabel_tersisa[i]]]
  }
  
  dimnames(A) <- list(NULL, variabel_tersisa)
  
  w <- matrix(unlist(weights_df$weight), ncol = 1, byrow = TRUE)
  n = 59
  v0 = 1
  alpha = 0.05
  t = 1
  
  Mean.Rt <- colMeans(A)
  Sigma.Rt <- cov(A)
  
  return.sim <- mvrnorm(n, Mean.Rt, Sigma.Rt)
  Rp <- return.sim %*% w
  
  rataRp <- mean(Rp)
  sd <- sd(Rp)
  Rbintang <- quantile(Rp, alpha)
  VaR <- v0 * Rbintang * sqrt(t)
  
  return(list(VaR = VaR, Mean.Rt = Mean.Rt, Sigma.Rt = Sigma.Rt, w = w,rataRp=rataRp,Rp=Rp,return.sim=return.sim))
}

# step 17 : MEMBANGKITKAN VALUE AT RISK SEBANYAK M KALI (AGAR STABIL, PEMBANGKITAN BERAKHIR JIKA SELISIH VAR SUDAH MENCAPAI 0,001)
generate_var_until_convergence <- function(n, Mean.Rt, Sigma.Rt, w, v0 = 1, alpha = 0.05, t = 1, targetError = 0.001) {
  # Fungsi untuk menghitung nilai X
  calculateX <- function(Rp) {
    return(v0 * quantile(Rp, p = alpha) * sqrt(t))
  }
  
  # Inisialisasi variabel
  Rp_prev <- matrix(0, nrow = n, ncol = 1)  # Rp sebelumnya
  Rp_current <- matrix(0, nrow = n, ncol = 1)  # Rp saat ini
  error <- Inf  # Inisialisasi selisih error
  X_values <- numeric(0)  # Variabel untuk menyimpan nilai X
  
  # Looping hingga selisih error kurang dari targetError
  while (error > targetError) {
    # Hitung Rp saat ini
    Rp_current <- mvrnorm(n, Mean.Rt, Sigma.Rt) %*% w
    
    # Hitung nilai X saat ini
    currentX <- calculateX(Rp_current)
    
    # Hitung selisih error
    if (length(X_values) > 0) {
      error <- abs(currentX - X_values[length(X_values)])
    }
    
    # Simpan nilai X
    X_values <- c(X_values, currentX)
    
    # Simpan Rp saat ini sebagai Rp sebelumnya untuk iterasi berikutnya
    Rp_prev <- Rp_current
    
    # Cetak nilai X dan selisih error
    cat("X:", currentX, " | Selisih Error:", error, "\n")
  }
  
  # Hitung rata-rata X
  averageX <- mean(X_values)
  
  # Cetak hasil
  cat("Rata-rata X:", averageX, "\n")
  
  return(list(X_values = X_values, averageX = averageX))
}

# Step 18: Expected Return of Portfolio
calculate_expected_return_portfolio <- function(weights_df, data, variabel_tersisa) {
  weights <- as.numeric(weights_df$weight)
  expected_returns <- numeric(length(variabel_tersisa))
  
  for (i in seq_along(variabel_tersisa)) {
    var <- variabel_tersisa[i]
    if (is.numeric(data[[var]])) {
      expected_returns[i] <- weights[i] * mean(data[[var]], na.rm = TRUE)
    }
  }
  
  expected_return_portfolio <- sum(expected_returns)
  return(expected_return_portfolio)
}

# Step 19: Calculate Portfolio Performance
calculate_portfolio_performance <- function(Rp, Rbr, alpha_beta_list, variabel_tersisa, weights_df) {
  # Hitung beta_portofolio di luar perulangan
  beta_portofolio <- sum(weights_df$weight * sapply(variabel_tersisa, function(v) alpha_beta_list[[paste0("beta_", v)]]))
  treynor <- (mean(Rp) - Rbr) / beta_portofolio
  return(list(treynor = treynor, beta_portofolio = beta_portofolio))
}

step_1 <- read_and_create_variables("D:/IGRADE/datareturn.xlsx")
# Step 2
variabel_tersisa <- calculate_means(step_1$data, step_1$Rbr)
# Step 3
step_3 <- calculate_residuals(variabel_tersisa$variabel_tersisa, step_1$data)

# Step 4
test_covariance_assumption(step_3$residuals_list)

# Step 5
test_covariance_ihsg<-test_covariance_with_ihsg(variabel_tersisa$variabel_tersisa, step_3$residuals_list, step_1$X)

# Step 6
alpha_beta_list <- calculate_alpha_beta(variabel_tersisa$variabel_tersisa, step_1$data)

# Step 7
erb_list <- calculate_erb(variabel_tersisa$variabel_tersisa, alpha_beta_list$alpha_beta_list, step_1$data)

# Step 8
nilai_erb_urut <- sort_erb(erb_list,variabel_tersisa$variabel_tersisa)

# Step 9
ai_bi_list <- calculate_ai_bi(nilai_erb_urut$variabel_tersisa, alpha_beta_list$alpha_beta_list, step_3$residuals_list, step_1$data, step_1$Rbr)

# Step 10
ci_list <- calculate_ci(nilai_erb_urut$variabel_tersisa, ai_bi_list$ai_list,ai_bi_list$bi_list, step_1$X)

# Step 11
sort_variabel_tersisa(nilai_erb_urut$variabel_tersisa, nilai_erb_urut$data_ERB,ai_bi_list$ai_list,ai_bi_list$bi_list, ci_list)

# Step 12
result_determine_vars <- determine_variables_in_portfolio(nilai_erb_urut$variabel_tersisa, ci_list, erb_list)
C_bintang <- determine_cut_off_point(ci_list, result_determine_vars$variabel_max_ci)

# Step 13
z_list <- calculate_z(result_determine_vars$variabel_tersisa, alpha_beta_list$alpha_beta_list, erb_list, ci_list, C_bintang, step_3$residuals_list)

# Step 14
weights_df <- calculate_weights(result_determine_vars$variabel_tersisa, z_list)

# Step 15
multivariate_normality<-calculate_mahalanobis_test(result_determine_vars$variabel_tersisa,step_1$data)
# Step 16
var_1_simulation <- calculate_var_1_simulation(step_1$data, result_determine_vars$variabel_tersisa, weights_df)

# Step 17
result_step_17 <- generate_var_until_convergence(n=57, Mean.Rt=var_1_simulation$Mean.Rt,
                                                 Sigma.Rt=var_1_simulation$Sigma.Rt, w=var_1_simulation$w)

# Step 18
expected_return_portfolio <- calculate_expected_return_portfolio(weights_df, step_1$data, result_determine_vars$variabel_tersisa)

# Step 19
treynor_values <- calculate_portfolio_performance(var_1_simulation$rataRp, step_1$Rbr, alpha_beta_list$alpha_beta_list, result_determine_vars$variabel_tersisa, weights_df)


# Main Function to Execute All Steps
execute_all_steps <- function(file_path) {
  # Step 1
  step_1 <- read_and_create_variables("D:/IGRADE/datareturn.xlsx")
  
  # Step 2
  variabel_tersisa <- calculate_means(step_1$data, step_1$Rbr)
  
  # Step 3
  step_3 <- calculate_residuals(variabel_tersisa$variabel_tersisa, step_1$data)
  
  # Step 4
  test_covariance_assumption(step_3$residuals_list)
  
  # Step 5
  test_covariance_ihsg<-test_covariance_with_ihsg(variabel_tersisa$variabel_tersisa, step_3$residuals_list, step_1$X)
  
  # Step 6
  alpha_beta_list <- calculate_alpha_beta(variabel_tersisa$variabel_tersisa, step_1$data)
  
  # Step 7
  erb_list <- calculate_erb(variabel_tersisa$variabel_tersisa, alpha_beta_list$alpha_beta_list, step_1$data)
  
  # Step 8
  nilai_erb_urut <- sort_erb(erb_list,variabel_tersisa$variabel_tersisa)
  
  # Step 9
  ai_bi_list <- calculate_ai_bi(nilai_erb_urut$variabel_tersisa, alpha_beta_list$alpha_beta_list, step_3$residuals_list, step_1$data, step_1$Rbr)
  
  # Step 10
  ci_list <- calculate_ci(nilai_erb_urut$variabel_tersisa, ai_bi_list$ai_list,ai_bi_list$bi_list, step_1$X)
  
  # Step 11
  sort_variabel_tersisa(nilai_erb_urut$variabel_tersisa, nilai_erb_urut$data_ERB,ai_bi_list$ai_list,ai_bi_list$bi_list, ci_list)
  
  # Step 12
  result_determine_vars <- determine_variables_in_portfolio(nilai_erb_urut$variabel_tersisa, ci_list, erb_list)
  C_bintang <- determine_cut_off_point(ci_list, result_determine_vars$variabel_max_ci)
  
  # Step 13
  z_list <- calculate_z(result_determine_vars$variabel_tersisa, alpha_beta_list$alpha_beta_list, erb_list, ci_list, C_bintang, step_3$residuals_list)
  
  # Step 14
  weights_df <- calculate_weights(result_determine_vars$variabel_tersisa, z_list)
  
  # Step 15
  multivariate_normality<-calculate_mahalanobis_test(result_determine_vars$variabel_tersisa,step_1$data)
  # Step 16
  var_1_simulation <- calculate_var_1_simulation(step_1$data, result_determine_vars$variabel_tersisa, weights_df)
  
  # Step 17
  result_step_17 <- generate_var_until_convergence(n=59, Mean.Rt=var_1_simulation$Mean.Rt,
                                                   Sigma.Rt=var_1_simulation$Sigma.Rt, w=var_1_simulation$w)
  
  # Step 18
  expected_return_portfolio <- calculate_expected_return_portfolio(weights_df, step_1$data, result_determine_vars$variabel_tersisa)
  
  # Step 19
  treynor_values <- calculate_portfolio_performance(var_1_simulation$rataRp, step_1$Rbr, alpha_beta_list$alpha_beta_list, result_determine_vars$variabel_tersisa, weights_df)
  
  # Return the results or use them as needed
  return(list(
    step_1 = step_1,
    variabel_tersisa = variabel_tersisa,
    step_3 = step_3,
    alpha_beta_list = alpha_beta_list,
    erb_list = erb_list,
    ai_bi_list = ai_bi_list,
    ci_list = ci_list,
    C_bintang = C_bintang,
    z_list = z_list,
    weights_df = weights_df,
    result_determine_vars,
    var_1_simulation = var_1_simulation,
    result_step_17=result_step_17,
    expected_return_portfolio = expected_return_portfolio,  # Fixed the variable name here
    treynor_values = treynor_values
  ))
}

result <- execute_all_steps("D:/IGRADE/datareturn.xlsx")

