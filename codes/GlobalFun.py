# 编译日期：2020-05-30 11:11:58
# 版权所有：www.i-search.com.cn
# coding=utf-8
import pandas as pd

# Load the Excel file
df = pd.read_excel('C:/Users/zhguo/Desktop/营销产品.xlsx')

# Assuming there is only one row of product data after the header
product_data = df.iloc[0]  # This gets the first row of data, skipping the header

# Construct the paragraph
paragraph = f"{df.columns[0]}: {product_data[0]}, "
for i in range(1, len(df.columns)):
    paragraph += f"{df.columns[i]}: {product_data[i]}, "

# Remove the last comma and space
paragraph = paragraph.strip(', ')

# Add the paragraph to a new column named 'Ad. Input'
df['Ad. Input'] = paragraph

# Save the dataframe back to the same Excel file, or to a new file if preferred
df.to_excel('C:/Users/zhguo/Desktop/营销产品.xlsx', index=False)

print("Paragraph added to the Excel file successfully.")