import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

df = pd.read_csv('input/sales.csv')
df.head()
df.info()
df.describe()

df['Date'] = pd.to_datetime(df['Date'])
monthly_sales = df.resample('M', on='Date').sum()
sales_mean = df['Total Amount'].mean() 
print(f"Avg Sales by month: {sales_mean:.2f}")

plt.plot(monthly_sales.index, monthly_sales['Total Amount'])
plt.title('Monthly Sales Trends')
plt.xlabel('Month')
plt.ylabel('Sales')
plt.show()

# 2️⃣ Create a "Month" column for the X axis
monthly_sales_reset = monthly_sales.reset_index()
monthly_sales_reset['Month'] = monthly_sales_reset['Date'].dt.strftime('%b')

# 3️⃣ Plot the barplot
plt.figure(figsize=(10,5))
sns.barplot(x='Month', y='Total Amount', data=monthly_sales_reset, color='skyblue', errorbar=None)
plt.title("Monthly Sales")
plt.xlabel("Month")
plt.ylabel("Total Sales Amount")
plt.show()