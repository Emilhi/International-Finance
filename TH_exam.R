
#===================================================================
#                            README
#
# Remove the "#" from line 12 to install a package control manager
# It will install and load all packages in the function "p_load"
#
# Change the variable "path" to where the dataset...
# "Data_MidTermExam_IntlFinanceV2.xlsx" is saved
#===================================================================

# install.packages("pacman")

library("pacman")
p_load(tidyverse, readxl, foreign, ggplot2, broom)

#========================= Load data =========================

path = "~/Example/path/to/data/file/Group/Data_MidTermExam_IntlFinanceV2.xlsx"

Q2a.FX_spot_rate <- read_excel(path,
                               sheet = "Q2a. FX spot rate",
                               range = "A2:C154")

Q2a.Bond_yields_3M <- read_excel(path,sheet = "Q2a. Bond rates",
                                 range = "A5:C157")

Q2a.Bond_yields_1Y <- read_excel(path,
                                 sheet = "Q2a. Bond rates",
                                 range = "F5:H157")

Q2a.Bond_yields_3Y <- read_excel(path,
                                 sheet = "Q2a. Bond rates",
                                 range = "K5:M157",
                                 col_types = c("text", "numeric",
                                               "numeric"))

Q2a.Bond_yields_5Y <- read_excel(path,
                                 sheet = "Q2a. Bond rates",
                                 range = "P5:R157",
                                 col_types = c("text", "numeric",
                                               "numeric"))

Q2a.Bond_yields_10Y <- read_excel(path,
                                  sheet = "Q2a. Bond rates",
                                  range = "U5:W157")

Q2b.Spot_rates <- read_excel(path,
                             sheet = "Q2b. Spot rates")

Q2b.Spot_rates <- Q2b.Spot_rates %>%
  rename(Date = me)

Q2b.Forward_rates <- read_excel(path,
                                sheet = "Q2b. Forward rates",
                                col_types = c("date",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric",
                                              "numeric", "numeric"
                                              )
                                )

Q2b.Forward_rates <- Q2b.Forward_rates %>%
  rename(Date = me)


#==============================================================
#                     PART B1
#==============================================================

# Renaming variables for spot rate
Q2a.FX_spot_rate <- Q2a.FX_spot_rate %>%
  rename("S_t(JPY/USD)" = `MSCI JPY TO 1 USD - EXCHANGE RATE`,
         "S_t(USD/JPY)" = `MSCI USD TO 1 JPY - EXCHANGE RATE`
         )

# Taking the log of spot rates
Q2a.FX_spot_rate <- Q2a.FX_spot_rate %>%
  mutate("ln_S_t(JPY/USD)" = log(`S_t(JPY/USD)`),
         "ln_S_t(USD/JPY)" = log(`S_t(USD/JPY)`)
         )

# Leading the log spot rates k-periods
k1 <- 1
k2 <- 4
k3 <- 4*3
k4 <- 4*5
k5 <- 4*10

Q2a.FX_spot_rate <- Q2a.FX_spot_rate %>%
  mutate("lead_3M_ln_S_t(USD/JPY)" = lead(`ln_S_t(USD/JPY)`, n = k1),
         "lead_1Y_ln_S_t(USD/JPY)" = lead(`ln_S_t(USD/JPY)`, n = k2),
         "lead_3Y_ln_S_t(USD/JPY)" = lead(`ln_S_t(USD/JPY)`, n = k3),
         "lead_5Y_ln_S_t(USD/JPY)" = lead(`ln_S_t(USD/JPY)`, n = k4),
         "lead_10Y_ln_S_t(USD/JPY)" = lead(`ln_S_t(USD/JPY)`, n = k5)
         )

# Taking differences  
Q2a.FX_spot_rate <- Q2a.FX_spot_rate %>%
  mutate("delta_s_(t+k1)" = `lead_3M_ln_S_t(USD/JPY)` - `ln_S_t(USD/JPY)`,
         "delta_s_(t+k2)" = `lead_1Y_ln_S_t(USD/JPY)` - `ln_S_t(USD/JPY)`,
         "delta_s_(t+k3)" = `lead_3Y_ln_S_t(USD/JPY)` - `ln_S_t(USD/JPY)`,
         "delta_s_(t+k4)" = `lead_5Y_ln_S_t(USD/JPY)` - `ln_S_t(USD/JPY)`,
         "delta_s_(t+k5)" = `lead_10Y_ln_S_t(USD/JPY)` - `ln_S_t(USD/JPY)`
         )

# Annualized differences
Q2a.FX_spot_rate <- Q2a.FX_spot_rate %>%
  mutate("annual_delta_s_(t+k1)" = 4*(`delta_s_(t+k1)`),
         "annual_delta_s_(t+k2)" = (`delta_s_(t+k2)`),
         "annual_delta_s_(t+k3)" = (`delta_s_(t+k3)`)/3,
         "annual_delta_s_(t+k4)" = (`delta_s_(t+k4)`)/5,
         "annual_delta_s_(t+k5)" = (`delta_s_(t+k5)`)/10
         )

# Prepare interest rates for regression
#drop na-values
drop_na_Q2a.Bond_yields_3M <- drop_na(Q2a.Bond_yields_3M)
drop_na_Q2a.Bond_yields_1Y <- drop_na(Q2a.Bond_yields_1Y)
drop_na_Q2a.Bond_yields_3Y <- drop_na(Q2a.Bond_yields_3Y)
drop_na_Q2a.Bond_yields_5Y <- drop_na(Q2a.Bond_yields_5Y)
drop_na_Q2a.Bond_yields_10Y <- drop_na(Q2a.Bond_yields_10Y)


# Take difference between home and foreing interest rate
drop_na_Q2a.Bond_yields_3M <- drop_na_Q2a.Bond_yields_3M %>%
  mutate("i_min_i*" = US - Japan)
drop_na_Q2a.Bond_yields_1Y <- drop_na_Q2a.Bond_yields_1Y %>%
  mutate("i_min_i*" = US - Japan)
drop_na_Q2a.Bond_yields_3Y <- drop_na_Q2a.Bond_yields_3Y %>%
  mutate("i_min_i*" = US - Japan)
drop_na_Q2a.Bond_yields_5Y <- drop_na_Q2a.Bond_yields_5Y %>%
  mutate("i_min_i*" = US - Japan)
drop_na_Q2a.Bond_yields_10Y <- drop_na_Q2a.Bond_yields_10Y %>%
  mutate("i_min_i*" = US - Japan)

# Create tables for regression

reg.3M <- Q2a.FX_spot_rate %>%
  select(c(Date, `annual_delta_s_(t+k1)`))
reg.3M <- inner_join(reg.3M,drop_na_Q2a.Bond_yields_3M) %>%
  select(-c(US, Japan))
reg.3M <- drop_na(reg.3M)

reg.1Y <- Q2a.FX_spot_rate %>%
  select(c(Date, `annual_delta_s_(t+k2)`))
reg.1Y <- inner_join(reg.1Y,drop_na_Q2a.Bond_yields_1Y) %>%
  select(-c(US, Japan))
reg.1Y <- drop_na(reg.1Y)

reg.3Y <- Q2a.FX_spot_rate %>%
  select(c(Date, `annual_delta_s_(t+k3)`))
reg.3Y <- inner_join(reg.3Y,drop_na_Q2a.Bond_yields_3Y) %>%
  select(-c(US, Japan))
reg.3Y <- drop_na(reg.3Y)

reg.5Y <- Q2a.FX_spot_rate %>%
  select(c(Date, `annual_delta_s_(t+k4)`))
reg.5Y <- inner_join(reg.5Y,drop_na_Q2a.Bond_yields_5Y) %>%
  select(-c(US, Japan))
reg.5Y <- drop_na(reg.5Y)

reg.10Y <- Q2a.FX_spot_rate %>%
  select(c(Date, `annual_delta_s_(t+k5)`))
reg.10Y <- inner_join(reg.10Y,drop_na_Q2a.Bond_yields_10Y) %>%
  select(-c(US, Japan))
reg.10Y <- drop_na(reg.10Y)

# Regressions
lm_reg.3M <- lm(`annual_delta_s_(t+k1)` ~ `i_min_i*`,data = reg.3M)
lm_reg.1Y <- lm(`annual_delta_s_(t+k2)` ~ `i_min_i*`, data = reg.1Y)
lm_reg.3Y <- lm(`annual_delta_s_(t+k3)` ~ `i_min_i*`, data = reg.3Y)
lm_reg.5Y <- lm(`annual_delta_s_(t+k4)` ~ `i_min_i*`, data = reg.5Y)
lm_reg.10Y <- lm(`annual_delta_s_(t+k5)` ~ `i_min_i*`, data = reg.10Y)


# Hypothesis testing: beta

# calculate p-value for t stat:  (slope_est - (1)) / std err

# k = 3M
slope_est <- lm_reg.3M$coefficients[2]
slope_se <-summary(lm_reg.3M)$coefficients[2,"Std. Error"] 
t_slope <- (slope_est - (1)) / slope_se
pval_neg1_3M <- 2*pt(-abs(t_slope),df=length(reg.3M$Date)-2)
names(pval_neg1_3M) <- NULL

# k = 1Y
slope_est <- lm_reg.1Y$coefficients[2]
slope_se <-summary(lm_reg.1Y)$coefficients[2,"Std. Error"] 
t_slope <- (slope_est - (1)) / slope_se
pval_neg1_1Y <- 2*pt(-abs(t_slope),df=length(reg.1Y$Date)-2)
names(pval_neg1_1Y) <- NULL

# k = 3Y
slope_est <- lm_reg.3Y$coefficients[2]
slope_se <-summary(lm_reg.3Y)$coefficients[2,"Std. Error"] 
t_slope <- (slope_est - (1)) / slope_se
pval_neg1_3Y <- 2*pt(-abs(t_slope),df=length(reg.3Y$Date)-2)
names(pval_neg1_3Y) <- NULL

# k = 5Y
slope_est <- lm_reg.5Y$coefficients[2]
slope_se <-summary(lm_reg.5Y)$coefficients[2,"Std. Error"] 
t_slope <- (slope_est - (1)) / slope_se
pval_neg1_5Y <- 2*pt(-abs(t_slope),df=length(reg.5Y$Date)-2)
names(pval_neg1_5Y) <- NULL

# k = 10Y
slope_est <- lm_reg.10Y$coefficients[2]
slope_se <-summary(lm_reg.10Y)$coefficients[2,"Std. Error"] 
t_slope <- (slope_est - (1)) / slope_se
pval_neg1_10Y <- 2*pt(-abs(t_slope),df=length(reg.10Y$Date)-2)
names(pval_neg1_10Y) <- NULL

# summarise pValues in a table
pvals_slope_equal_1 <- tibble()
pvals_slope_equal_1[1,1] = "pValue"
pvals_slope_equal_1[1,2] = pval_neg1_3M
pvals_slope_equal_1[1,3] = pval_neg1_1Y
pvals_slope_equal_1[1,4] = pval_neg1_3Y
pvals_slope_equal_1[1,5] = pval_neg1_5Y
pvals_slope_equal_1[1,6] = pval_neg1_10Y

pval_names <- c("Variable","3M", "1Y", "3Y", "5Y", "10Y")

pvals_slope_equal_1 <- pvals_slope_equal_1 %>%
  rename_at(vars(starts_with(".")), ~ pval_names)





#==============================================================
#                     PART B2
#==============================================================

# Create log forward discount
dates <- Q2b.Forward_rates %>%
  select(Date)

no_date_fwd <- Q2b.Forward_rates %>%
  select(-Date)

no_date_spot <- Q2b.Spot_rates %>%
  select(-Date)

fwd_discount <- log(no_date_fwd) - log(no_date_spot)

fwd_discount <- cbind(dates, fwd_discount)


replace_na_fwd_discount <- fwd_discount %>%
  gather(units_exchng_rate, discount, -Date) %>%
  replace_na(list(discount = 0)) %>%
  group_by(Date)%>%
  arrange(discount, .by_group = T) %>%
  ungroup()

Portfolio_nbrs <- as_tibble(rep(1:6, times = 336, each = 7))

Portfolio_nbrs <- Portfolio_nbrs %>%
  rename(Portfolio_nbr = value)

replace_na_fwd_discount <- cbind(replace_na_fwd_discount,Portfolio_nbrs)

unit_portfolio_date <- replace_na_fwd_discount

ln_spot_rates_long <- Q2b.Spot_rates %>%
  gather(units_exchng_rate, spot, -Date) %>%
  mutate(ln_spot = log(spot)) %>%
  select(-spot)

ln_spot_rates_long_t1 <- ln_spot_rates_long %>%
  group_by(units_exchng_rate) %>%
  mutate(ln_spot_t1 = lead(ln_spot, order_by = units_exchng_rate)) %>%
  ungroup() %>%
  select(-ln_spot)

ln_fwd_rates_long <- Q2b.Forward_rates %>%
  gather(units_exchng_rate, forward, -Date) %>%
  mutate(ln_forward = log(forward)) %>%
  select(-forward)

ln_fwd_and_spot_t1_rates_long <- cbind(ln_fwd_rates_long, ln_spot_rates_long_t1$ln_spot_t1)

ln_fwd_and_spot_t1_rates_long <- ln_fwd_and_spot_t1_rates_long %>%
  rename(ln_spot_t1 = `ln_spot_rates_long_t1$ln_spot_t1`)

unit_portfolio_date <- left_join(unit_portfolio_date, ln_fwd_and_spot_t1_rates_long)

unit_portfolio_date <- unit_portfolio_date %>%
  mutate("log_excess_ret" = ln_forward - ln_spot_t1)

# Add date identifier
dates <- dates %>%
  mutate(date_id = seq(1:336))

unit_portfolio_date <- right_join(dates, unit_portfolio_date, by = "Date")

#initialize df's to store values
avg_discounts <- as_tibble(dates$Date) %>%
  rename(Date = value)

avg_log_excess_ret <- as_tibble(dates$Date) %>%
  rename(Date = value)

# Replace Na values from excess return
unit_portfolio_date <- unit_portfolio_date %>%
  replace_na(list(log_excess_ret = 0))


# Find average forward discount rate,
# and average log currency excess return

for (k in 1:6) {
  for (j in 1:336) {
    
    port_filter <- unit_portfolio_date %>%
      filter(Portfolio_nbr == k,
             date_id == j)
    
    avg_discounts[j,k+1] = sapply(port_filter[,4], mean, na.rm = F)
    avg_log_excess_ret[j,k+1] = sapply(port_filter[,8], mean, na.rm = F)
    
  }
}

new_names <- c("p1","p2","p3","p4","p5","p6")

avg_discounts <- avg_discounts %>%
  rename_at(vars(starts_with(".")), ~ new_names)
avg_log_excess_ret <- avg_log_excess_ret %>%
  rename_at(vars(starts_with(".")), ~ new_names)

avg_discount.sum <- avg_discounts %>%
  summarise(p1 = 12*mean(p1, na.rm = T),
            p2 = 12*mean(p2, na.rm = T),
            p3 = 12*mean(p3, na.rm = T),
            p4 = 12*mean(p4, na.rm = T),
            p5 = 12*mean(p5, na.rm = T),
            p6 = 12*mean(p6, na.rm = T))


avg_excess_ret.sum <- avg_log_excess_ret %>%
  summarise(p1 = 12*mean(p1, na.rm = T),
            p2 = 12*mean(p2, na.rm = T),
            p3 = 12*mean(p3, na.rm = T),
            p4 = 12*mean(p4, na.rm = T),
            p5 = 12*mean(p5, na.rm = T),
            p6 = 12*mean(p6, na.rm = T))


avg_excess_ret.sum.sd <- avg_log_excess_ret %>%
  summarise(p1 = 12*sd(p1, na.rm = T),
            p2 = 12*sd(p2, na.rm = T),
            p3 = 12*sd(p3, na.rm = T),
            p4 = 12*sd(p4, na.rm = T),
            p5 = 12*sd(p5, na.rm = T),
            p6 = 12*sd(p6, na.rm = T))


avg_discount.sum <- avg_discount.sum %>%
  pivot_longer(cols = starts_with("p"),
               names_to = "Portfolio",
               values_to = "Avg. forward discount")

avg_excess_ret.sum <- avg_excess_ret.sum %>%
  pivot_longer(cols = starts_with("p"),
               names_to = "Portfolio",
               values_to = "Avg. excess return (E[r_p])")

avg_excess_ret.sum.sd <- avg_excess_ret.sum.sd %>%
  pivot_longer(cols = starts_with("p"),
               names_to = "Portfolio",
               values_to = "A Standard deviation")

sharpe_ratio <- avg_excess_ret.sum
sharpe_ratio <- left_join(sharpe_ratio, avg_excess_ret.sum.sd)
sharpe_ratio <- sharpe_ratio %>%
  mutate("A Sharpe" = (`Avg. excess return (E[r_p])`/`A Standard deviation`)*sqrt(12)) %>%
  select(-c(`Avg. excess return (E[r_p])`, `A Standard deviation`))

output <- left_join(avg_discount.sum, avg_excess_ret.sum)
output <- left_join(output, avg_excess_ret.sum.sd)
output <- left_join(output, sharpe_ratio)

output <- output %>%
  pivot_longer(cols = starts_with("A"),
               names_to = "Variable",
               values_to = "Value")
output <- output %>%
  pivot_wider(names_from = "Portfolio",
              values_from = "Value")

output[3,1] = "Standard deviation"
output[4,1] = "Sharpe ratio"


# Crate a long-short portfolio (portf. 2-6 on 1)
ls_portfolio <- avg_log_excess_ret %>%
  mutate(p2_min_p1 = p2-p1,
         p3_min_p1 = p3-p1,
         p4_min_p1 = p4-p1,
         p5_min_p1 = p5-p1,
         p6_min_p1 = p6-p1) %>%
  select(-c(p1,p2,p3,p4,p5,p6))

ls_portfolio.sum <- ls_portfolio %>%
  summarise(p2_min_p1 = 12*mean(p2_min_p1, na.rm = T),
            p3_min_p1 = 12*mean(p3_min_p1, na.rm = T),
            p4_min_p1 = 12*mean(p4_min_p1, na.rm = T),
            p5_min_p1 = 12*mean(p5_min_p1, na.rm = T),
            p6_min_p1 = 12*mean(p6_min_p1, na.rm = T))


ls_portfolio.sd <- ls_portfolio %>%
  summarise(p2_min_p1 = 12*sd(p2_min_p1, na.rm = T),
            p3_min_p1 = 12*sd(p3_min_p1, na.rm = T),
            p4_min_p1 = 12*sd(p4_min_p1, na.rm = T),
            p5_min_p1 = 12*sd(p5_min_p1, na.rm = T),
            p6_min_p1 = 12*sd(p6_min_p1, na.rm = T))

ls_portfolio.sum <- ls_portfolio.sum %>%
  pivot_longer(cols = starts_with("p"),
               names_to = "Portfolio",
               values_to = "Avg. excess return")
ls_portfolio.sd <- ls_portfolio.sd %>%
  pivot_longer(cols = starts_with("p"),
               names_to = "Portfolio",
               values_to = "A Standard deviation")

ls_portfolio.SR <- ls_portfolio.sum
ls_portfolio.SR <- left_join(ls_portfolio.SR, ls_portfolio.sd)
ls_portfolio.SR <- ls_portfolio.SR %>%
  mutate("A SR" = (`Avg. excess return`/`A Standard deviation`)*sqrt(12))

ls_portfolio.SR <- ls_portfolio.SR %>%
  pivot_longer(cols = starts_with("A"),
               names_to = "Variable",
               values_to = "Value")
ls_portfolio.SR <- ls_portfolio.SR %>%
  pivot_wider(names_from = "Portfolio",
              values_from = "Value")

ls_portfolio.SR[2,1] = "Standard deviation"
ls_portfolio.SR[3,1] = "Sharpe ratio"

# Q. B6
#=======

# Define confidence level for two sided test
conf_level <- 0.01/2

# Find mean excess return from p1 and p6
mean_excess_ret_p1 <- as.numeric(output[2,2])
mean_excess_ret_p6 <- as.numeric(output[2,7])

# Difference between the log currency excess return of
# the last and the first portfolios
mean_excess_ret_p6_min_p1 <- mean_excess_ret_p6 - mean_excess_ret_p1

log_excess_ret_p6_min_p1 <- avg_log_excess_ret %>%
  mutate(p6_min_p1 = p6 - p1) %>%
  select(-c(p1, p2, p3, p4, p5, p6))

log_excess_ret_p6_min_p1.sd <- log_excess_ret_p6_min_p1 %>%
  summarise("Standard deviation" = 12*sd(p6_min_p1, na.rm = T))

log_excess_ret_p6_min_p1.se <- as.numeric(
  log_excess_ret_p6_min_p1.sd[1,1]
  )/sqrt(length(log_excess_ret_p6_min_p1$Date))


# Find t-stat for hypothesis test
t_stat_diff_means <- mean_excess_ret_p6_min_p1/log_excess_ret_p6_min_p1.se

# find critical value
crit_val_test_means <- abs(qt(conf_level, length(log_excess_ret_p6_min_p1$Date)-1))



# Q. B7
#=======

# answered in B6


#======================================================================
#                         PRINT ANSWERS
#======================================================================

print("==================PART B1====================")
print("Regression outputs:")
print("k = 3 months")
tidy(lm_reg.3M)
print("k = 1 year")
tidy(lm_reg.1Y)
print("k = 3 year")
tidy(lm_reg.3Y)
print("k = 5 year")
tidy(lm_reg.5Y)
print("k = 10 year")
tidy(lm_reg.10Y)

print("Hypothesis tests: H0: beta = 1")
pvals_slope_equal_1

print("==================PART B2====================")

print("2B: Main computations:")
output
ls_portfolio.SR

print(c("2B.Q7:", mean_excess_ret_p6_min_p1))
print("2B.Q6:")
print(c("t-stat = ", t_stat_diff_means))
print(c("Critical value = ", crit_val_test_means))
print("We reject the null hypothesis on all levels, since |t-stat|> critical value")


