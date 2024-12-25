import yfinance as yf
import matplotlib.pyplot as plt

# Fetch VIX data
vix_data = yf.download('^VIX', start='2023-12-01', end='2024-12-31')
# vix_data = yf.download('^VIX', start='2004-12-01', end='2024-12-31')

# Extract the dates and VIX closing values
dates = vix_data.index
vix_close = vix_data['Close']

# Plot the VIX data
plt.figure(figsize=(12, 6))
plt.plot(dates, vix_close, label='VIX Close', color='gray', linewidth=1.0)

# Highlight regions where VIX is above 30
# plt.fill_between(dates, vix_close, 30, where=(vix_close > 30), color='red', alpha=0.5, label='Above 30')
_vix_close = vix_close.values.flatten()
threshold = 20
# threshold = 40
plt.fill_between(dates, _vix_close, threshold,
                 where=(_vix_close > threshold),
                 color='red', alpha=0.5, label=f'Above {threshold}')

# Add threshold lines
plt.axhline(y=threshold, color='red', linestyle='--', label=f'Threshold ({threshold} pts)')
# plt.axhline(y=30, color='red', linestyle='--', label='Threshold (30 pts)')
# plt.axhline(y=20, color='orange', linestyle='--', label='Threshold (20 pts)')

# Add labels and title
plt.title(f'VIX Index with Highlights Above {threshold}', fontsize=16)
plt.xlabel('Date', fontsize=12)
plt.ylabel('VIX Level', fontsize=12)
plt.legend(fontsize=10)
plt.grid(True)

# Show the plot
plt.show()
