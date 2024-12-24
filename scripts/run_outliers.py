
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import yfinance as yf
import numpy as np
from scipy import stats

# Step 1: Download historical data for market and stock_
market = 'SPY'
stock = 'INTC'
# stock = 'NVDA'
# stock = 'AMD'

# stock = 'DPZ'
# stock = 'TSM'
#
# stock = 'QBTS'
# stock = 'LAES'
#
# stock = 'BIDU'
# stock = 'BABA'
# stock = 'JD'
# stock = 'PDD'
#
# ETF
# stock = 'QQQ'
# stock = 'MAGS'
# stock = 'SOXX'

tickers = [market, stock]
data = yf.download(tickers, start='2023-12-10', end='2024-12-10')
# data = yf.download(tickers, start='2019-12-10', end='2024-12-10')
print("market {} and stock {}".format(market, stock))

# Step 2: Calculate daily returns
returns = data['Adj Close'].pct_change().dropna()

# Step 3: Perform linear regression to find slope and intercept
slope, intercept, r_value, p_value, std_err = stats.linregress(returns[market], returns[stock])
print("slope {:.2f} intercept {:.2f} r_value {:.2f} p_value {:.2f} stderr {:.2f}".format(
    slope, intercept, r_value, p_value, std_err))
print(f"Slope of the regression line: {slope:.2f}")

# Step 4: Identify outliers based on z-scores
# threshold = 5  # Define a threshold for outliers
threshold = 4
# threshold = 3.5
z_scores = np.abs(stats.zscore(np.column_stack((returns[market], returns[stock]))))
outliers = (z_scores > threshold).any(axis=1)

# Step 5: Create scatter plot with regression line and highlight outliers
plt.figure(figsize=(10, 6))

# Scatter plot for normal points
plt.scatter(returns[market][~outliers], returns[stock][~outliers], color='gray', label='Normal Points', alpha=0.7)

# Scatter plot for outliers
plt.scatter(returns[market][outliers], returns[stock][outliers], color='red', label='Outliers', s=100)  # Larger size for visibility

# Plotting the regression line
plt.plot(returns[market], slope * returns[market] + intercept, color='green', linewidth=2, label='Regression Line')

# Step 6: Center the view on the plot
# plt.xlim(-0.1, 0.1)  # Adjust these values based on your data range
# plt.ylim(-0.1, 0.1)  # Adjust these values based on your data range
plt.xlim(returns[market].min() - 0.01, returns[market].max() + 0.01)
plt.ylim(returns[stock].min() - 0.01, returns[stock].max() + 0.01)

plt.title('1-Year Daily Returns: {} vs {} with Outliers Highlighted'.format(market, stock))
plt.xlabel(f'{market} Daily Returns')
plt.ylabel(f'{stock} Daily Returns')
plt.grid(True)

# Add horizontal and vertical lines at zero for reference
plt.axhline(0, color='gray', lw=0.8, ls='--')
plt.axvline(0, color='gray', lw=0.8, ls='--')

# Step 7: Print details of outlier points including volume
outlier_dates = returns.index[outliers]
outlier_market_returns = returns[market][outliers]
outlier_stock_returns = returns[stock][outliers]
stock_prices = data['Adj Close'][stock][outlier_dates]
stock_volumes = data['Volume'][stock][outlier_dates]

print("Average volume: {:.2f}m".format(data['Volume'][stock].mean() / 1e6))

if len(outliers[outliers]) > 0:
    print("\nFound {} outliers".format(len(outliers[outliers])))
    print("Outlier Details:")
    for (date, market_return,
         stock_return, stock_price, volume) in zip(outlier_dates, outlier_market_returns,
                                                   outlier_stock_returns, stock_prices, stock_volumes):
        print(f"Date: {date.date()},"
              f" {market} Return: {market_return*100:.2f} %, {stock} Return: {stock_return * 100:.2f} %,"
              f" {stock} Price: {stock_price:.2f}, Volume: {volume / 1e6:.2f}m")
else:
    print("No outlier was found")

# Show legend
plt.legend()

# Show plot
plt.show()
