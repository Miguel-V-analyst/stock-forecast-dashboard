#R code that produces three time series plots for the below stock symbols. 

#=================================

rm(list = ls())
gc()
options(stringsAsFactors = FALSE)

library(openxlsx)
library(tidyquant)
library(tidyverse)
library(timetk)
library(lubridate)
library(forecast)

# ========= Styles =========
money_style  <- createStyle(numFmt = "$#,##0.00")
pct_style    <- createStyle(numFmt = "0.00%")
date_style   <- createStyle(numFmt = "m/d/yyyy")
title_style  <- createStyle(textDecoration = "bold", fontSize = 14)
subtle_style <- createStyle(fontColour = "#6B7280")

# ========= Get data =========
tickers <- c("AAPL","GOOG","NFLX","NVDA")

stock_data_table <- tickers %>%
  tq_get(from = "2020-01-01", to = Sys.Date())

stock_pivot_table <- stock_data_table %>%
  pivot_table(
    .rows    = ~ YEAR(date),
    .columns = ~ symbol,
    .values  = ~ PCT_CHANGE_FIRSTLAST(adjusted)
  ) %>%
  rename(year = 1)

# ========= Forecasting (ARIMA on log price) =========
forecast_days <- 90

pred_all <- stock_data_table %>%
  select(symbol, date, adjusted) %>%
  group_by(symbol) %>%
  group_modify(~{
    y <- log(.x$adjusted)
    fit <- auto.arima(y)
    fc  <- forecast(fit, h = forecast_days)
    
    tibble(
      date = seq(max(.x$date) + 1, by = "day", length.out = forecast_days),
      pred = exp(as.numeric(fc$mean)),
      lo80 = exp(as.numeric(fc$lower[,"80%"])),
      hi80 = exp(as.numeric(fc$upper[,"80%"])),
      lo95 = exp(as.numeric(fc$lower[,"95%"])),
      hi95 = exp(as.numeric(fc$upper[,"95%"]))
    )
  }) %>%
  ungroup()

# ========= Decision Dashboard =========
horizon_days <- 30

last_prices <- stock_data_table %>%
  group_by(symbol) %>%
  summarise(last_date = max(date), last_actual = last(adjusted), .groups = "drop")

forecast_30 <- pred_all %>%
  group_by(symbol) %>%
  slice(horizon_days) %>%
  ungroup() %>%
  rename(forecast_date = date,
         pred_30 = pred,
         lo80_30 = lo80,
         hi80_30 = hi80)

dashboard_tbl <- last_prices %>%
  inner_join(forecast_30, by = "symbol") %>%
  mutate(
    exp_pct_change = (pred_30 / last_actual) - 1,
    signal = case_when(
      lo80_30 > last_actual ~ "Bullish",
      hi80_30 < last_actual ~ "Bearish",
      TRUE ~ "Neutral"
    )
  ) %>%
  select(symbol, last_date, last_actual,
         forecast_date, pred_30, exp_pct_change,
         lo80_30, hi80_30, signal)

# ========= Plot =========
stock_plot <- stock_data_table %>%
  group_by(symbol) %>%
  plot_time_series(date, adjusted, .color_var = symbol, .facet_ncol = 2, .interactive = FALSE)

# ========= Forecast wide table =========
pred_wide <- pred_all %>%
  mutate(date = as.Date(date)) %>%
  pivot_wider(
    names_from = symbol,
    values_from = c(pred, lo80, hi80, lo95, hi95),
    names_glue = "{symbol}_{.value}"
  ) %>%
  arrange(date)

# ========= Workbook =========
wb <- createWorkbook()

# --- Decision Dashboard ---
addWorksheet(wb, "Decision_Dashboard", gridLines = FALSE)
writeData(wb, "Decision_Dashboard", "Decision Dashboard", startRow = 1, startCol = 1)
addStyle(wb, "Decision_Dashboard", title_style, rows = 1, cols = 1)
writeData(wb, "Decision_Dashboard",
          paste("Last refresh:", Sys.time(), "| Horizon:", horizon_days, "days"),
          startRow = 2, startCol = 1)
addStyle(wb, "Decision_Dashboard", subtle_style, rows = 2, cols = 1)

writeDataTable(wb, "Decision_Dashboard", dashboard_tbl, startRow = 4, startCol = 1,
               tableStyle = "TableStyleMedium9", withFilter = TRUE)

n <- nrow(dashboard_tbl)
data_rows <- 5:(5 + n - 1)
addStyle(wb, "Decision_Dashboard", date_style,  rows = data_rows, cols = c(2,4), gridExpand = TRUE, stack = TRUE)
addStyle(wb, "Decision_Dashboard", money_style, rows = data_rows, cols = c(3,5,7,8), gridExpand = TRUE, stack = TRUE)
addStyle(wb, "Decision_Dashboard", pct_style,   rows = data_rows, cols = 6, gridExpand = TRUE, stack = TRUE)
setColWidths(wb, "Decision_Dashboard", cols = 1:ncol(dashboard_tbl), widths = "auto")
freezePane(wb, "Decision_Dashboard", firstRow = TRUE)

# --- Stock Forecast Analysis ---
addWorksheet(wb, "Stock_forecast_analysis", gridLines = FALSE)
writeData(wb, "Stock_forecast_analysis", "Stock Forecasting Analysis", startRow = 1, startCol = 1)
addStyle(wb, "Stock_forecast_analysis", title_style, rows = 1, cols = 1)
writeData(wb, "Stock_forecast_analysis", paste("Last refresh:", Sys.time()), startRow = 2, startCol = 1)

writeDataTable(wb, "Stock_forecast_analysis", stock_pivot_table, startRow = 4, startCol = 1,
               tableStyle = "TableStyleMedium9", withFilter = TRUE)

addStyle(wb, "Stock_forecast_analysis", pct_style,
         rows = 5:(5 + nrow(stock_pivot_table) - 1),
         cols = 2:ncol(stock_pivot_table),
         gridExpand = TRUE, stack = TRUE)

setColWidths(wb, "Stock_forecast_analysis", cols = 1:ncol(stock_pivot_table), widths = "auto")
freezePane(wb, "Stock_forecast_analysis", firstRow = TRUE)

print(stock_plot)
insertPlot(wb, "Stock_forecast_analysis", startCol = 7, startRow = 4, width = 8, height = 4)  

# --- Forecast 90d ---
addWorksheet(wb, "Forecast_90d", gridLines = FALSE)
writeData(wb, "Forecast_90d", "90-Day Forecast (ARIMA on log price)", startRow = 1, startCol = 1)
addStyle(wb, "Forecast_90d", title_style, rows = 1, cols = 1)

writeDataTable(wb, "Forecast_90d", pred_wide, startRow = 3, startCol = 1,
               tableStyle = "TableStyleMedium2", withFilter = TRUE)

addStyle(wb, "Forecast_90d", date_style,
         rows = 4:(4 + nrow(pred_wide) - 1), cols = 1,
         gridExpand = TRUE, stack = TRUE)

addStyle(wb, "Forecast_90d", money_style,
         rows = 4:(4 + nrow(pred_wide) - 1),
         cols = 2:ncol(pred_wide),
         gridExpand = TRUE, stack = TRUE)

setColWidths(wb, "Forecast_90d", cols = 1:ncol(pred_wide), widths = "auto")
freezePane(wb, "Forecast_90d", firstRow = TRUE)


desired_order <- c("Decision_Dashboard", "Stock_forecast_analysis", "Forecast_90d")
worksheetOrder(wb) <- match(desired_order, names(wb))


path <- "/Users/miguelvelez/Desktop"
saveWorkbook(wb, file = file.path(path, "StockForecastAnalysis.xlsx"), overwrite = TRUE)
openXL(file.path(path, "StockForecastAnalysis.xlsx"))
