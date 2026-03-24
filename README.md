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



## 7. Sales Entry System

To make daily sales recording easier, I built a **Sales Entry Form** instead of entering transactions directly into the Sales table. This form acts as the interface where the business owner records customer orders in a structured way.

![Sales Entry Form](images/sales_entry_form.png)

*Sales Entry Form used to record customer transactions.*

### Automated Fields

Some fields in the form are generated automatically to simplify data entry.

- **Date** is automatically generated. Each time the form is cleared, the system updates the date to the current day.
- **Sales ID** is automatically generated to uniquely identify each order recorded in the system.

This helps the business owner keep track of how many orders have been recorded.

---

### Product Selection Logic

The form uses dropdown lists linked to the **Products table** to ensure that only valid product combinations can be selected.

The selection follows the product hierarchy:

Category → Brand → Product Name

For example, when the user selects **Body Spray** as the category, the brand dropdown will only display brands that belong to that category. After selecting a brand, the product dropdown shows only products under that brand.

Once a product is selected, the system automatically retrieves the following information from the Products table using **XLOOKUP formulas**:

- Product ID  
- Product Size  
- Cost Price  
- Selling Price  

This ensures that product information remains consistent across the system.

---

### Recording a Sales Transaction

After selecting the product, the user enters:

- Customer Name  
- Quantity purchased  
- Sales Channel

Once the **Submit** button is clicked, the transaction is automatically written into the **Sales table**, where it becomes part of the system’s transaction records.

---

### Form Control Buttons

The form includes three buttons that control how transactions are handled.

**Submit**

Records the sales transaction and sends the data to the Sales table using VBA automation.

**Clear**

Resets the form after a transaction has been recorded, preparing it for the next customer order. Clearing the form also generates the next Sales ID.

**Delete**

Removes the most recent transaction from the Sales table. This is useful when a transaction is recorded incorrectly or entered twice.

---

### Why the Entry Form Matters

Without the entry form, the business owner would need to manually type transactions directly into the Sales table. This increases the risk of errors and makes the system harder to manage.

By introducing a structured entry form with automated fields and dropdown selections, the system ensures that sales records are captured consistently while keeping the transaction process fast and simple.


## 8. Purchase Entry System

In addition to recording sales, the system also includes a **Purchase Entry Form** used to record inventory restocking. This form allows the business owner to track products purchased from vendors and update stock levels in the system.

![Purchase Entry Form](images/purchase_entry_form.png)

*Purchase Entry Form used to record inventory restocking transactions.*

### Purpose of the Purchase Entry Form

The purpose of this form is to ensure that every product restocked from the market is properly recorded. Instead of typing purchase records directly into the Purchases table, the business owner records transactions through this form.

This approach keeps the purchase records structured and prevents accidental edits to the transaction table.

---

### Product Selection

Just like the Sales Entry Form, the Purchase Entry Form uses dropdown selections linked to the **Products table**.

The user selects:

Category → Brand → Product Name

Once the product is selected, the system automatically retrieves key product information such as:

- Product ID  
- Product Size  
- Cost Price  

This ensures that the correct product details are always recorded during restocking.

---

### Recording a Purchase Transaction

To record a purchase transaction, the user enters:

- Vendor Name  
- Quantity Purchased  
- Purchase Date

After entering these details, the user clicks the **Submit** button. The system then automatically writes the purchase record into the **Purchases table**.

Each purchase transaction increases the available stock of that product in the system.

---

### Form Control Buttons

The Purchase Entry Form also includes control buttons to simplify the recording process.

**Submit**

Records the purchase transaction and transfers the data into the Purchases table using VBA automation.

**Clear**

Resets the form after a purchase has been recorded, allowing the user to quickly record another restocking transaction.

---

### Why This Form Matters

Without a structured purchase recording system, it becomes difficult to track how much inventory has been restocked over time.
By recording all restocking transactions through the Purchase Entry Form, the system ensures that inventory inflow is captured correctly. This data is later used by the Inventory table to calculate current stock levels and inventory value.


## 9. Product Price Update Module

Product prices in a retail business can change over time, especially when restocking products from vendors. Instead of manually searching through the Products table to update prices, the system includes a **Product Price Update Module** that allows price changes to be handled quickly and safely.

![Product Price Update Form](images/product_price_update_form.png)

*Product Price Update module used to modify product cost and selling prices.*

### Purpose of the Update Module

The purpose of this module is to make it easy for the business owner to update product prices whenever market prices change.

When a product is restocked and the cost price increases, the business owner can adjust both the **Cost Price** and **Selling Price** using this form without editing the Products table directly.

This prevents accidental modifications to the product data and keeps price updates organized.

---

### Product Selection Process

The update form follows the same product hierarchy used throughout the system:

Category → Brand → Product Name

The user first selects the product category, then the brand, and finally the specific product. Once the product is selected, the system automatically retrieves the following information from the **Products table**:

- Product ID  
- Product Size  
- Current Cost Price  
- Current Selling Price  

These values appear automatically in the form using **XLOOKUP formulas**.

---

### Updating Product Prices

After the product details are retrieved, the user enters:

- New Cost Price  
- New Selling Price  

When the **Update button** is clicked, the system automatically updates the corresponding row in the **Products table**.

This ensures that all future sales transactions use the updated price.

---

### Why This Module Matters

Without a dedicated update system, price changes would require manually editing rows in the Products table. This can easily lead to mistakes, especially when managing many products.

By introducing a controlled update form, the system ensures that product pricing remains accurate while keeping the data management process simple for the business owner.


## 10. Automation Implemented with VBA

To make the system easier to use and reduce repetitive manual work, I implemented several automations using VBA (Visual Basic for Applications).

![VBA Automation Example](images/vba_code_example.png)


These automations allow the forms and dashboard to interact directly with the system tables and data model, ensuring that transactions are recorded correctly while keeping the workflow simple for the business owner.

### Submit Automation

The Submit buttons in both the Sales Entry Form and Purchase Entry Form are powered by VBA.

When the Submit button is clicked, the VBA script automatically transfers the data entered in the form into the appropriate transaction table.

For example:

- Sales transactions are written to the Sales table  
- Purchase transactions are written to the Purchases table  

This removes the need for manual copying or editing of tables.

---

### Clear Form Automation

The Clear button resets the form fields after a transaction has been recorded.

This prepares the form for the next customer order and automatically generates the next Sales ID, allowing the business owner to continue recording transactions without interruption.

---

### Delete Automation

The Delete button allows the user to remove the most recent transaction from the Sales table.

This feature is useful when a transaction is entered incorrectly or submitted twice. Instead of manually locating the row in the table, the user can simply click the Delete button to remove the entry.

---

### Product Price Update Automation

The Update button in the Product Price Update Module is also powered by VBA.

When the user enters a new cost price and selling price and clicks the Update button, the VBA script locates the correct row in the Products table and updates the pricing information automatically.

---

### Dashboard Refresh Automation

The dashboard includes a Refresh button powered by VBA.

After recording sales or purchase transactions, the dashboard does not update automatically because Pivot Tables and the Data Model require a refresh to display the latest data.

When the Refresh button is clicked, the VBA script refreshes:

- Pivot Tables  
- Data Model connections  
- Dashboard charts and KPIs  

This allows the business owner to immediately see updated business performance after recording transactions.


