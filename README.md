# perfume-inventory-management-system
Excel-based inventory management system for a perfume business that tracks sales, purchases, stock levels, and profitability. Built using Power Pivot, DAX, and VBA automation.


## Project Overview

This project builds a structured inventory management system in Microsoft Excel for a perfume retail business. The system allows the business owner to record daily sales transactions, track inventory restocking, manage product pricing, and monitor business performance through an interactive dashboard.

Many small retail businesses rely on notebooks or scattered spreadsheets to manage inventory and sales. This often leads to inaccurate stock records, difficulty identifying profitable products, and limited visibility into overall business performance. This system was designed to address those challenges by providing a structured way to capture operational data and transform it into useful insights.

The solution is built around multiple connected tables within the Excel Data Model, including **Products**, **Sales**, **Purchases**, and **Inventory** tables. Sales and purchase entry forms capture daily transactions, which feed into the data model where **DAX measures** calculate key business metrics such as revenue, profit, stock levels, and inventory value. **VBA automation** is used to simplify data entry, manage form actions, and refresh analytical outputs.

Together, these components create a workflow where daily business activities are recorded, structured, and automatically translated into analytical insights. The dashboard then provides a clear overview of sales performance, inventory status, product profitability, and stock risks, helping the business owner make more informed operational and restocking decisions.


## 3. Business Problem

Many small perfume retailers manage their inventory and sales using notebooks, memory, or simple spreadsheets. While this may work when the business is small, it quickly becomes unreliable as the number of products and daily transactions increases.

In practice, many vendors end up relying on guesswork to make decisions. For example, a perfume seller might assume that a particular fragrance has finished because customers ask for it often, without having accurate records showing how many units were actually sold. Similarly, restocking decisions are sometimes made based on intuition rather than real sales data. This makes it difficult to know which perfumes are truly performing well and which ones are not generating enough profit.

Another challenge is the lack of clear visibility into product profitability. A business owner may see that sales are happening but still struggle to understand which products are generating the most revenue, which ones carry higher costs, and how much profit the business is actually making.

As the product catalog grows across multiple **categories, brands, and product sizes**, tracking inventory manually becomes even more complicated. Without a structured system, it becomes difficult to answer important operational questions such as:

- Which perfumes are selling the most?
- Which products generate the highest profit?
- How much stock is currently available for each item?
- When should a product be restocked?

This project was developed to address these challenges by creating a structured system where product data, sales transactions, and purchase records are captured in a centralized environment. By organizing the data into connected tables and analytical models, the system replaces guesswork with reliable insights that help the business owner track inventory accurately, understand product performance, and make better operational decisions.



## 4. Tools and Technologies Used

The system was implemented using **Microsoft Excel with Power Pivot for data modeling, DAX for analytical calculations, and VBA for automation**. These components work together to transform raw transaction records into a structured inventory tracking and business analytics system.


### Microsoft Excel

Microsoft Excel serves as the core platform for the entire system. It is used to design the user interface for data entry, store structured business data, manage product records, and present analytical insights through a dashboard. Excel provides the environment where operational workflows, data modeling, and reporting are integrated into a single solution.

### Excel Tables

Structured Excel Tables are used to store and organize the core business data within the system. These include the **Products Table, Sales Table, Purchase Table, and Inventory Table**. Using structured tables ensures consistent data formatting, improves formula reliability, and allows the data model to reference transactional records dynamically as the system grows.

### Excel Data Model (Power Pivot)

The **Excel Data Model**, powered by Power Pivot, is used to connect the multiple tables that form the foundation of the system. Relationships are established between the Products table and the Sales, Purchases, and Inventory tables, creating a relational structure inside Excel. This allows data from separate operational tables to be analyzed together without duplicating information.

### DAX (Data Analysis Expressions)

DAX measures are used to perform analytical calculations within the data model. These calculations generate key business metrics such as **Total Revenue, Cost of Goods Sold, Total Profit, Profit Margin, Current Stock Quantity, and Inventory Value**. By placing these calculations in the data model, the system separates raw transaction data from analytical logic, enabling dynamic reporting through the dashboard.

### VBA (Visual Basic for Applications)

VBA is used to automate several operational tasks within the system. Macros control actions performed through the entry forms, including submitting sales transactions, recording purchase entries, clearing form inputs, deleting incorrect records, and refreshing the analytical dashboard. This automation improves usability and reduces the need for manual data manipulation.

### Pivot Tables and Pivot Charts

Pivot Tables and Pivot Charts are used as the reporting interface connected to the Excel Data Model. Instead of calculating metrics directly inside pivot tables, the system relies on DAX measures from the data model. Pivot tables therefore serve as the visualization layer that aggregates and displays the calculated metrics through the dashboard.

### Data Validation

Data Validation is used within the entry forms to create controlled dropdown selections for product attributes such as **category, brand, and product name**. This ensures that data entered into the system follows a consistent structure and prevents incorrect or inconsistent product entries. It also supports the hierarchical product structure used in the system, where product selection follows the relationship between category, brand, and individual product names.


## 5. System Architecture

This project was designed as a structured Excel system rather than a single spreadsheet. Instead of storing everything in one sheet, the system separates product information, transactions, inventory calculations, and analysis into different connected tables.

The goal of this structure is simple: business data should be **recorded once, stored in the correct table, and then used to generate insights automatically**.

The system is built around **four core data tables** supported by entry forms, automation, and a dashboard.

Core tables in the system:

- Products Table  
- Sales Table  
- Purchase Table  
- Inventory Table  

These tables work together to capture business operations and feed the analytical dashboard.

---

### Products Table

The **Products table** stores all the information about the products sold in the business. Every product in the system is defined here before it can be used in sales or purchase records.

Columns in this table include:

- Product ID  
- Product Name  
- Brand  
- Category  
- Size (ML)  
- Cost Price  
- Selling Price  

This table acts as the **main reference table** for the entire system. Sales and purchase transactions both depend on the information stored here.

*Screenshot: Products Table*

---

### Sales Table

The **Sales table** records every product sold to customers. Data in this table comes directly from the **Sales Entry Form**.

Each row represents one transaction and includes information such as:

- Sales ID  
- Product Name  
- Quantity Sold  
- Sales Channel  
- Date of Transaction  

This table is used to calculate important business metrics such as revenue, profit, and product performance.

*Screenshot: Sales Table*

---

### Purchase Table

The **Purchase table** records inventory restocking activities. When the business buys products from suppliers, the transaction is recorded through the **Purchase Entry Form**.

Each row contains information such as:

- Vendor Name  
- Product Purchased  
- Quantity Purchased  
- Purchase Date  

These records allow the system to track how much stock enters the business.

*Screenshot: Purchase Table*

---

### Inventory Table

The **Inventory table** tracks the current stock level for every product.

It combines information from the purchase table and sales table to calculate:

- Total Quantity Purchased  
- Total Quantity Sold  
- Current Stock Quantity  

The table also calculates the **current value of inventory**, which shows how much money is currently tied up in stock.

This table acts as the bridge between operational data and the analytical dashboard.

*Screenshot: Inventory Table*

---

### Entry Forms

To make the system easier to use, the project includes two entry forms.

**Sales Entry Form**

Allows the user to record sales transactions using dropdown selections instead of typing directly into the sales table.

**Purchase Entry Form**

Allows the user to record restocking transactions by selecting the product, vendor, and quantity purchased.

These forms help reduce data entry mistakes and keep transaction records consistent.

*Screenshot: Sales Entry Form*  
*Screenshot: Purchase Entry Form*

---

### Price Update Module

The system also includes a **Product Price Update form**. This allows the business owner to update product prices when market prices change.

Instead of editing the product table manually, the user can:

1. Select the product  
2. Enter the new cost price  
3. Enter the new selling price  
4. Click the update button  

The system automatically updates the product table.

*Screenshot: Price Update Form*

---

### Analytical Layer

All the operational tables feed into the **Excel Data Model using Power Pivot**. The data model allows analytical calculations to be written using **DAX measures**.

These calculations power the dashboard and track business metrics such as:

- Total Revenue  
- Total Profit  
- Profit Margin  
- Inventory Value  
- Sales Performance  
- Product Performance  

*Screenshot: Dashboard*

---

### System Workflow

The system works in the following sequence:

1. Product information is stored in the Products table  
2. Sales are recorded through the Sales Entry Form  
3. Restocking transactions are recorded through the Purchase Entry Form  
4. Inventory levels are calculated automatically  
5. The Power Pivot data model processes the data  
6. The dashboard displays business insights  

This structure allows the system to function as a **simple inventory management and business analysis tool built entirely in Excel**.


## 6. Data Model Structure

Before building the dashboard, I had to first structure the data properly. One thing I learned while working on this project is that dashboards become easy only after the data behind them is organized correctly. At the beginning, some of my pivot tables were not behaving the way I expected. I kept trying different fixes until I realized the real issue was not the pivot tables. The real issue was how the tables were structured and connected.

To solve this, I moved the tables into the **Excel Data Model** and created proper relationships between them. Once the tables were connected correctly, the system started behaving more like a small relational database instead of disconnected spreadsheets.

The system is built around several structured tables that work together.

---

### Products Table (Master Product Table)

The **Products table** acts as the master reference table for every item sold in the store. Every product is registered here before it can appear in any sale or purchase transaction.

This table contains the core product information, including:

- Product ID  
- Product Name  
- Brand  
- Category  
- Size  
- Cost Price  
- Selling Price  

Each product has a **unique Product ID**, which acts as the key that connects product information to other tables in the system.

Another important structure in this table is the **product hierarchy**:

Category → Brand → Product Name

For example:

Body Spray → Riggs → Patrol  
Body Spray → Smart Collection → No.06  

This hierarchy makes the dashboard filters work properly. When a user selects a category like *Body Spray*, the dashboard automatically limits the brands and products shown to only those that belong to that category.

![Products Table](images/products_table.png)

*Products table showing product hierarchy (Category → Brand → Product Name).*

---

### Sales Table (Transaction Table)

The **Sales table** stores every sales transaction recorded through the Sales Entry Form. Each row represents a customer purchase.

Fields stored in this table include:

- Date  
- Category  
- Brand  
- Product Name  
- Size  
- Product ID  
- Customer Name  
- Quantity  
- Unit Price  
- Total Amount  
- Sales ID  
- Sales Channel  

The **Product ID** links each sale back to the Products table so the system always knows which product was sold.

This table acts as one of the **transaction tables** in the system because it records business activity happening over time.

---

### Purchases Table (Inventory Restocking)

The **Purchases table** records inventory coming into the store when products are restocked from vendors.

This table is populated through the Purchase Entry Form and contains:

- Date  
- Category  
- Brand  
- Product Name  
- Size  
- Product ID  
- Vendor Name  
- Quantity Purchased  
- Unit Cost  
- Total Cost  

Each purchase increases the available stock of that product in the business.

Like the Sales table, this table also acts as a **transaction table** because it records operational activity.

---

### Inventory Table (Stock Monitoring Table)

The **Inventory table** provides a consolidated view of the stock position for every product in the store. It combines information from the Purchases and Sales tables to calculate how much stock is currently available.

Important fields in this table include:

- Product ID  
- Category  
- Brand  
- Product Name  
- Size  
- Cost Price  
- Total Purchased  
- Total Sold  
- Current Stock  
- Stock Status  

The **Current Stock** value is calculated using the relationship between purchases and sales.

Current Stock = Total Purchased − Total Sold

Based on this result, the system assigns a **stock status** such as:

- In Stock  
- Low Stock  
- Out of Stock  

This allows the business owner to quickly identify products that need restocking and avoid running out of popular items.

![Inventory Table](images/inventory_table.png)

*Inventory table calculating current stock using purchase and sales data.*

---

### Table Relationships in the Data Model

All tables are connected inside the **Excel Data Model** using the **Product ID** field.

The relationships are structured as follows:

Products (Product ID) → Sales (Product ID)  
Products (Product ID) → Purchases (Product ID)

The **Products table acts as the central reference table**, while the Sales and Purchases tables store transaction data connected to those products.

In simple terms:

- The Products table tells the system **what products exist**  
- The Purchases table records **products entering the business**  
- The Sales table records **products leaving the business**

Because these tables are connected inside the Data Model, Excel can combine information from multiple tables when calculating metrics or building reports.

Without this structure, pivot tables would struggle to properly combine sales, purchases, and product data across multiple sheets.

![Power Pivot Relationships](images/data_model_relationships.png)

*Power Pivot Diagram View showing relationships between Products, Sales, and Purchases tables.*

---

### Why the Data Model Matters

Once the tables were properly structured and connected, the system became much more powerful.

The Data Model allows the system to:

- combine data from multiple tables
- calculate business metrics across transactions
- support DAX calculations
- drive dashboard visualizations

This structure also made it possible to create analytical measures such as **Total Revenue, Profit, Inventory Value, and Stock Risk indicators**, which power the dashboard analysis shown later in this project.
