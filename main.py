import pandas as pd

# Read CSV files into DataFrames
product_df = pd.read_csv('data/product.csv', sep=';')
special_offer_df = pd.read_csv('data/special_offer.csv', sep=';')
product_category_df = pd.read_csv('data/product_category.csv', sep=';')
product_subcategory_df = pd.read_csv('data/product_subcategory.csv', sep=';')
sales_df = pd.read_csv('data/sales.csv', sep=';')
sales_details_df = pd.read_csv('data/sales_details.csv', sep=';')

# Create a dict to store data for the "Tables" sheet
tables_data = {
    "Table Name": ["product", "product_category", "product_subcategory", "sales", "special_offer", "sales_details"],
    "Description": ["Information about products", "Information about product categories",
                    "Information about product subcategories", "Information about sales transactions",
                    "Information about special offers", "Details of sales transactions"],

    "Total Records": [
        len(product_df),
        len(product_category_df),
        len(product_subcategory_df),
        len(sales_df),
        len(special_offer_df),
        len(sales_details_df)
    ],

    "Data Granularity": [
        "Product level", "Category level",
        "Subcategory level", "Transaction level",
        "Offer level",
        "Transaction detail level"
    ],

    "Business Logic": [
        "Various attributes of products",
        "Hierarchical categorization of products",
        "Further categorization within categories",
        "Sales data including date, customer, and product information",
        "Details of special offers",
        "Detailed information about each sale"
    ],

    "Business Questions": [
        "Which products are top sellers?",
        "How are products categorized?",
        "What are the subcategories within each category?",
        "What are our sales trends over time?",
        "How effective are our special offers?",
        "What are the details of individual sales transactions?"
    ]
}

# Create a DataFrame for the "Tables" sheet
tables_df = pd.DataFrame(tables_data)

# Create a new Excel file and write the "Tables" sheet
with pd.ExcelWriter("Data Dictionary.xlsx") as writer:
    tables_df.to_excel(writer, sheet_name="Tables", index=False)

# Write "Table Details" DataFrame to Excel
excel_file = "Data Dictionary.xlsx"

# Define table names and corresponding column names
table_columns = {
    "product": product_df.columns.tolist(),
    "product_category": product_category_df.columns.tolist(),
    "product_subcategory": product_subcategory_df.columns.tolist(),
    "sales": sales_df.columns.tolist(),
    "special_offer": special_offer_df.columns.tolist(),
    "sales_details": sales_details_df.columns.tolist()
}

# Create a DataFrame to store table details
table_details_data = []

# Iterate over each table and its columns
for table, columns in table_columns.items():
    for column in columns:
        table_details_data.append([table, column])

# Create DataFrame for "Table Details" sheet
table_details_df = pd.DataFrame(table_details_data, columns=["Table Name", "Column Name"])

try:
    # Write "Table Details" DataFrame to Excel
    with pd.ExcelWriter(excel_file, mode='a', engine='openpyxl') as writer:
        # Check if "Table Details" sheet already exists
        if 'Table Details' in writer.book.sheetnames:
            # If sheet exists, delete it
            idx = writer.book.sheetnames.index('Table Details')
            writer.book.remove(writer.book.worksheets[idx])
            writer.book.save(excel_file)

        table_details_df.to_excel(writer, sheet_name="Table Details", index=False)

        # Fill in descriptions for special_offer table
        special_offer_desc = {
            "SpecialOfferID": "Unique identifier for each special offer",
            "Description": "Description of the special offer",
            "DiscountPct": "Percentage discount offered",
            "Type": "Type of special offer (e.g., discount, promotion)",
            "Category": "Category of the special offer"
        }
        for index, row in table_details_df.iterrows():
            if row["Table Name"] == "special_offer":
                table_details_df.at[index, "Description"] = special_offer_desc.get(row["Column Name"], "")

        table_details_df.to_excel(writer, sheet_name="Table Details", index=False)

except Exception as e:
    print("An error occurred:", str(e))


# Group sales_details by product and sum the quantities sold
product_sales = sales_details_df.groupby('ProductID')['OrderQty'].sum().reset_index()

# Merge with product_df to get product names
product_sales = pd.merge(product_sales, product_df[['ProductID', 'Name']], on='ProductID', how='left')

# Sort by quantity sold in descending order
product_sales = product_sales.sort_values(by='OrderQty', ascending=False)

# Print product names and total sold quantity in descending order
print("Product Names and Total Sold Quantity in Descending Order:")
print(product_sales[['Name', 'OrderQty']])

# Find the product with the highest quantity sold
highest_quantity_product = product_sales.iloc[0]
print("\nProduct with the Highest Quantity Sold:")
print("Name:", highest_quantity_product['Name'])
print("Quantity Sold:", highest_quantity_product['OrderQty'])

# Find the product with the lowest quantity sold
lowest_quantity_product = product_sales.iloc[-1]
print("\nProduct with the Lowest Quantity Sold:")
print("Name:", lowest_quantity_product['Name'])
print("Quantity Sold:", lowest_quantity_product['OrderQty'])

# Merge sales_details_df with special_offer_df to identify which products were on special offer
sales_details_with_offer = pd.merge(sales_details_df, special_offer_df, on='SpecialOfferID', how='left')

# Group by ProductID and count the number of orders
product_sales_with_offer = sales_details_with_offer.groupby('ProductID')['SalesOrderID'].nunique().reset_index()

# Merge with product_df to get product names
product_sales_with_offer = pd.merge(product_sales_with_offer, product_df[['ProductID', 'Name']], on='ProductID', how='left')

# Sort by number of orders in descending order
product_sales_with_offer = product_sales_with_offer.sort_values(by='SalesOrderID', ascending=False)

# Print product names and number of orders when on special offer
print("Product Names and Number of Orders When on Special Offer:")
print(product_sales_with_offer[['Name', 'SalesOrderID']])

# Find the product with the highest number of sales when on special offer
highest_sales_product = product_sales_with_offer.iloc[0]
print("\nProduct with the Highest Number of Sales When on Special Offer:")
print("Name:", highest_sales_product['Name'])
print("Number of Orders:", highest_sales_product['SalesOrderID'])

# Calculate unit discount and total discount
sales_details_df['unit_discount'] = sales_details_df['UnitPrice'] * sales_details_df['UnitPriceDiscount']
sales_details_df['total_discount'] = sales_details_df['OrderQty'] * sales_details_df['unit_discount']

# Calculate original price and discounted price
sales_details_df['original_price'] = sales_details_df['UnitPrice'] * sales_details_df['OrderQty']
sales_details_df['discounted_price'] = sales_details_df['original_price'] - sales_details_df['total_discount']

# Print the original price and discounted price for each product order
print()
print("Original Price and Discounted Price for Each Product Order:")
print(sales_details_df[['ProductID', 'OrderQty', 'original_price', 'discounted_price']])

# Group sales by CustomerID and count the number of unique orders
customer_orders = sales_df.groupby('CustomerID')['SalesOrderID'].nunique().reset_index()

# Find the maximum number of orders
max_orders = customer_orders['SalesOrderID'].max()

# Filter for customers with the maximum number of orders
most_orders_customers = customer_orders[customer_orders['SalesOrderID'] == max_orders]

# Print the ID(s) of the customer(s) with the most orders
print()
print("ID(s) of the Customer(s) with the Most Orders:")
print(most_orders_customers['CustomerID'].tolist())

# Calculate frequency segment
customer_orders['frequency'] = pd.cut(customer_orders['SalesOrderID'], bins=[0, 1, 3, float('inf')], labels=['New', 'Repeated', 'Fan'])

# Define monetary value bins and labels
monetary_bins = [0, 100, 10000, float('inf')]
monetary_labels = ['Frugal Spender', 'Medium Spender', 'High Spender']

# Calculate total purchase amount for each customer
customer_purchase_amount = sales_df.groupby('CustomerID')['TotalDue'].sum().reset_index()

# Calculate monetary value segment
customer_purchase_amount['monetary_value'] = pd.cut(customer_purchase_amount['TotalDue'], bins=monetary_bins, labels=monetary_labels)

# Merge customer_orders and customer_purchase_amount
customer_segments = pd.merge(customer_orders, customer_purchase_amount, on='CustomerID', how='outer')

# Create a matrix showing the number of customers belonging to each combination of segments
segment_matrix = pd.crosstab(customer_segments['frequency'], customer_segments['monetary_value'])

# Identify the best and worst customers based on the matrix
best_customers = segment_matrix.idxmax().max()
worst_customers = segment_matrix.idxmin().min()

# Print the matrix and best/worst customers
print("Matrix showing the number of customers belonging to each combination of segments:")
print(segment_matrix)
print("\nBest Customers (Based on Highest Segment):", best_customers)
print("Worst Customers (Based on Lowest Segment):", worst_customers)

# Calculate Revenue
total_revenue = sales_df['TotalDue'].sum()

# Calculate Sales Growth
previous_period_sales = sales_df['TotalDue'].shift(1)
sales_growth = ((sales_df['TotalDue'] - previous_period_sales) / previous_period_sales) * 100
average_sales_growth = sales_growth.mean()

# Calculate Customer Acquisition Cost (CAC)
# Assuming CAC is calculated based on marketing and sales expenses
# CAC = Total marketing and sales expenses / Number of new customers acquired
total_marketing_sales_expenses = 10000  # Example value for total marketing and sales expenses
number_of_new_customers = sales_df['CustomerID'].nunique()  # Assuming each customer ID represents a unique customer
cac = total_marketing_sales_expenses / number_of_new_customers

# Create a DataFrame for the KPIs
kpi_data = {
    'KPI': ['Revenue', 'Sales Growth', 'Customer Acquisition Cost (CAC)'],
    'Value': [total_revenue, average_sales_growth, cac]
}

kpi_df = pd.DataFrame(kpi_data)

# Export the DataFrame to an Excel file named "Insights"
kpi_df.to_excel('Insights.xlsx', index=False)

print("KPIs report has been successfully generated and saved as Insights.xlsx")
