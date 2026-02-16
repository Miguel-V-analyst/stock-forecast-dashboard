# Automated Stock Forecasting & Decision Dashboard (R)

## Overview

This project builds an automated financial analytics report using R.
It ingests live market data, forecasts future stock prices using time-series modeling, and generates a formatted Excel dashboard designed for decision-making rather than raw analysis.

The report refreshes automatically and produces a multi-sheet workbook containing historical performance, future projections, and an interpretable trading signal.

## Features

* Pulls daily stock price data via `tidyquant`
* Calculates yearly returns by company
* Fits ARIMA time-series models for each ticker
* Forecasts 90-day future price ranges with confidence intervals
* Generates a 30-day decision signal (Bullish / Neutral / Bearish)
* Builds a formatted Excel report using `openxlsx`
* Includes visualization and executive-style dashboard

## Technologies Used

* R
* tidyverse
* tidyquant
* forecast (ARIMA modeling)
* openxlsx
* timetk

## Output

The script automatically produces a 3-tab Excel report:

1. **Decision Dashboard** – summary and signal interpretation
2. **Stock Forecast Analysis** – yearly returns and visualizations
3. **Forecast 90d** – detailed predicted price ranges

## How to Run

1. Install required packages
2. Run `stock_forecast.R`
3. The Excel report will be generated in the specified output directory

## Purpose

This project demonstrates:

* Time-series forecasting
* Automated reporting
* Data engineering workflows
* Translating analytics into decision support

  ## Preview
### Decision Dashboard
![Decision Dashboard](images/decision_dashboard.png)

## How to Run
1. Open `stock_forecast.R`
2. Update the `path <- "..."` variable to a valid folder on your machine
3. Run the script in RStudio
4. The Excel workbook will be saved as `StockForecastAnalysis.xlsx`


### 90-Day Forecast
![90-Day Forecast](images/90_day_forecast.png)

