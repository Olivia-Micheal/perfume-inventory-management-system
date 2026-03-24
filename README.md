# perfume-inventory-management-system
Excel-based inventory management system for a perfume business that tracks sales, purchases, stock levels, and profitability. Built using Power Pivot, DAX, and VBA automation


## Project Overview

This project builds a structured inventory management system in Microsoft Excel for a perfume retail business. The system allows the business owner to record daily sales transactions, track inventory restocking, manage product pricing, and monitor business performance through an interactive dashboard.

Many small retail businesses rely on notebooks or scattered spreadsheets to manage inventory and sales. This often leads to inaccurate stock records, difficulty identifying profitable products, and limited visibility into overall business performance. This system was designed to address those challenges by providing a structured way to capture operational data and transform it into useful insights.

The solution is built around multiple connected tables within the Excel Data Model, including **Products**, **Sales**, **Purchases**, and **Inventory** tables. Sales and purchase entry forms capture daily transactions, which feed into the data model where **DAX measures** calculate key business metrics such as revenue, profit, stock levels, and inventory value. **VBA automation** is used to simplify data entry, manage form actions, and refresh analytical outputs.

Together, these components create a workflow where daily business activities are recorded, structured, and automatically translated into analytical insights. The dashboard then provides a clear overview of sales performance, inventory status, product profitability, and stock risks, helping the business owner make more informed operational and restocking decisions.


##  Business Problem

Many small perfume retailers manage their inventory and sales using notebooks, memory, or simple spreadsheets. While this may work when the business is small, it quickly becomes unreliable as the number of products and daily transactions increases.

In practice, many vendors end up relying on guesswork to make decisions. For example, a perfume seller might assume that a particular fragrance has finished because customers ask for it often, without having accurate records showing how many units were actually sold. Similarly, restocking decisions are sometimes made based on intuition rather than real sales data. This makes it difficult to know which perfumes are truly performing well and which ones are not generating enough profit.

Another challenge is the lack of clear visibility into product profitability. A business owner may see that sales are happening but still struggle to understand which products are generating the most revenue, which ones carry higher costs, and how much profit the business is actually making.

As the product catalog grows across multiple **categories, brands, and product sizes**, tracking inventory manually becomes even more complicated. Without a structured system, it becomes difficult to answer important operational questions such as:

- Which perfumes are selling the most?
- Which products generate the highest profit?
- How much stock is currently available for each item?
- When should a product be restocked?

This project was developed to address these challenges by creating a structured system where product data, sales transactions, and purchase records are captured in a centralized environment. By organizing the data into connected tables and analytical models, the system replaces guesswork with reliable insights that help the business owner track inventory accurately, understand product performance, and make better operational decisions.



##  Tools and Technologies Used

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


## System Architecture

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


##  Data Model Structure

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



##  Sales Entry System

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


##  Purchase Entry System

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


##  Product Price Update Module

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


##  Automation Implemented with VBA

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


##  Analytical Measures Built with DAX

To generate business insights from the data model, I created several analytical measures using **DAX (Data Analysis Expressions)** in Power Pivot. These measures power the KPI cards and charts displayed on the dashboard.

Because the tables are connected inside the Data Model, DAX allows the system to calculate metrics across multiple related tables such as **Sales, Purchases, Products, and Inventory**.

![DAX Measures](images/dax_measures.png)

*Example of DAX measures created in Power Pivot.*

---

### Total Revenue

Total Revenue represents the total amount of money generated from all sales transactions recorded in the system.

```
Total Revenue = SUM(Sales[Total Amount])
```

This metric helps the business owner quickly see the overall sales performance of the store.

---

### Cost of Goods Sold (COGS)

COGS represents the total cost of the products that have been sold.

It is calculated by multiplying the cost price of each product by the quantity sold.

```
COGS = SUMX(Sales, Sales[Quantity] * Sales[Cost Price])
```

Tracking COGS allows the system to determine how much it actually cost the business to generate the recorded sales.

---

### Total Profit

Total Profit measures how much money the business earned after subtracting the cost of goods sold from total revenue.

```
Total Profit = [Total Revenue] - [COGS]
```

This KPI helps the business owner understand whether the store is generating profit from its sales activities.

---

### Profit Margin

Profit Margin shows the percentage of revenue that remains as profit after accounting for product costs.

```
Profit Margin = DIVIDE([Total Profit], [Total Revenue])
```

This metric helps evaluate how efficiently the business converts sales into profit.

---

### Total Orders

Total Orders counts the number of sales transactions recorded in the system.

Because each order has a unique **Sales ID**, the system can count the number of completed transactions.

```
Total Orders = DISTINCTCOUNT(Sales[Sales ID])
```

This helps the business owner track the number of orders processed over time.

---

### Current Stock Quantity

Current Stock Quantity shows the number of product units currently available in inventory.

This value is calculated using the Inventory table.

```
Current Stock = Total Purchased - Total Sold
```

This metric helps the business owner monitor stock availability and avoid running out of products.

---

### Cost of Goods Available (COGA)

COGA represents the total value of inventory currently available in the store.

It is calculated by multiplying the **cost price of each product** by the **current stock quantity**.

```
COGA = SUMX(Inventory, Inventory[Current Stock] * Inventory[Cost Price])
```

This KPI helps the business owner understand how much capital is currently tied up in inventory.

---

### Stock Risk Count

Stock Risk Count identifies how many products are at risk of running out of stock.

Products are flagged based on their **stock status**, which can be:

- Low Stock  
- Out of Stock  

The measure counts the number of products whose stock status falls into either of these categories. This allows the business owner to quickly identify items that require immediate restocking.

---

Together, these measures transform raw transaction data into meaningful business insights. They allow the business owner to monitor sales performance, profitability, and inventory health from a single dashboard.


##  Dashboard and Business Insights

The dashboard serves as the main monitoring interface of the system. It combines information from the **Sales, Purchases, Products, and Inventory tables** and presents the most important business metrics in a simple visual format.

Instead of manually reviewing multiple tables to understand sales performance or stock levels, the dashboard provides a **central overview of the business in one place**.

![Dashboard Overview](images/dashboard_overview.png)

*Main dashboard showing key metrics, product performance, and inventory indicators.*

---

### Key Performance Indicators (KPIs)

At the top of the dashboard, several KPI cards summarize the most important business metrics.

These KPIs include:

- **Total Revenue** – total sales generated from recorded transactions.
- **Cost of Goods Sold (COGS)** – total cost of products sold.
- **Total Profit** – earnings after subtracting product costs from revenue.
- **Profit Margin** – percentage of revenue retained as profit.
- **Total Orders** – number of sales transactions recorded.
- **Current Stock Quantity** – number of product units currently available in inventory.
- **Cost of Goods Available (COGA)** – total value of inventory currently in stock.
- **Stock Risk Count** – number of products that are either low in stock or out of stock.

These KPIs provide a quick summary of both **sales performance and inventory health**.

---

### Revenue Performance Over Time

The dashboard includes a chart that tracks **revenue performance by month**.

![Revenue Trend](images/revenue_trend.png)

This visualization helps the business owner quickly identify:

- periods of strong sales
- periods where sales decline
- overall revenue trends across time

Tracking revenue trends makes it easier to understand customer demand and business growth patterns.

---

### Most Profitable Products

The dashboard highlights **top-performing products based on profit contribution**.

![Top Products by Profit](images/top_products_profit.png)

This helps the business owner quickly identify products that generate the most profit. These insights can support decisions such as:

- prioritizing restocking of high-performing products
- focusing promotions on profitable items
- identifying products that contribute the most to revenue growth

---

### Sales Channel Performance

The dashboard also analyzes **sales performance by channel**.

![Sales Channel Performance](images/sales_channel_performance.png)

This allows the business owner to see where sales are coming from and understand which channels contribute the most revenue.

---

### Inventory Monitoring

The system continuously monitors inventory by combining information from the **Sales and Purchases tables**.

This allows the dashboard to highlight:

- available inventory levels
- products that are running low
- products that are out of stock

These insights help the business owner restock products before inventory shortages occur.

---

### Interactive Filters

The dashboard includes several interactive filters that allow the user to explore the data from different perspectives.

Available filters include:

- **Category Filter**
- **Brand Filter**
- **Date Timeline**

For example:

If the user selects the **Body Spray** category, the dashboard automatically updates to show only the performance of body spray products.

The user can then filter further by **Brand**, such as selecting **Smart Collection**, to analyze the performance of that specific brand within the category.

This filtering capability allows the business owner to quickly answer questions such as:

- Which category generates the most revenue?
- Which brand performs best?
- How are specific products performing over time?

---

### Dashboard Refresh Button

Because the dashboard relies on pivot tables and the Excel Data Model, a **Refresh Button** powered by VBA was added to the dashboard.

![Refresh Button](images/dashboard_refresh_button.png)

After recording new sales or purchase transactions, the business owner can click the refresh button to update the entire model so that all charts and KPIs reflect the most recent data.

---

Overall, the dashboard transforms raw transaction data into actionable business insights, allowing the business owner to monitor **sales performance, profitability, and inventory health from a single interface**.



##  System Workflow

This section explains how data moves through the system from **daily transactions to business insights**.

The system was designed so the business owner interacts with simple entry forms while the calculations, relationships, and reporting happen automatically in the background.

---

### Step 1 — Product Setup

Every product must first be registered in the **Products Table** before it can be used anywhere in the system.

This table stores the master information for each item, including the product category, brand, product name, size, cost price, and selling price.

Each product is assigned a **unique Product ID**, which acts as the primary key used to connect product information across the system.

![Products Table](images/products_table.png)

---

### Step 2 — Recording Sales

Customer purchases are recorded using the **Sales Entry Form** rather than entering data directly into the sales table.

Once a sale is submitted, the system automatically writes the transaction into the **Sales Table**, including information such as the product, quantity sold, price, sales channel, and transaction ID.

Several fields are automatically generated or retrieved from the Products table, including:

- Product ID  
- Product Size  
- Product Price  
- Transaction Date  
- Sales ID  

This automation helps reduce manual errors and ensures consistent transaction records.

![Sales Entry Form](images/sales_entry_form.png)

---

### Step 3 — Recording Purchases

When the business owner restocks inventory from suppliers, the transaction is recorded through the **Purchase Entry Form**.

The submitted data is written to the **Purchases Table**, which tracks incoming inventory quantities and vendor information.

This allows the system to track how much inventory is entering the business over time.

![Purchase Entry Form](images/purchase_entry_form.png)

---

### Step 4 — Inventory Calculation

The system calculates the current stock level for each product by combining information from the **Sales Table** and the **Purchases Table**.

The basic logic used is:

Current Stock = Total Purchased − Total Sold

These calculations populate the **Inventory Table**, which continuously tracks the available stock level for every product.

![Inventory Table](images/inventory_table.png)

---

### Step 5 — Product Price Updates

The system includes a **Product Price Update Module** that allows the business owner to adjust product prices when supplier costs change.

When a new price is entered, the system updates the relevant fields in the **Products Table**, ensuring that future transactions reflect the updated cost and selling price.

![Product Update Module](images/product_update_module.png)

---

### Step 6 — Data Model Processing

Once transactions are recorded, the system processes the data using the **Excel Data Model**.

Relationships connect the core tables:

- Products  
- Sales  
- Purchases  
- Inventory  

These relationships allow Excel to combine information across multiple tables and support analytical calculations using **DAX measures**.

---

### Step 7 — Dashboard Insights

The processed data is then visualized through the **dashboard**, which displays key business metrics, product performance indicators, and inventory status.

Interactive filters allow the user to analyze performance by category, brand, or time period.

![Dashboard Overview](images/dashboard_overview.png)

---

### Step 8 — System Refresh

After recording new sales or purchases, the business owner can click the **Refresh Button** on the dashboard.

This triggers a VBA macro that refreshes the entire model so that all KPIs, charts, and reports reflect the most recent data.

---

Through this workflow, the system converts **daily operational transactions into structured data and actionable business insights**, allowing the business owner to track sales performance, manage inventory efficiently, and make informed business decisions.



##  Skills Demonstrated and System Usage

This project required combining several technical and analytical skills to design a structured system that supports real business operations. The system was built not just as a spreadsheet, but as a small data-driven application that allows a business owner to manage sales, inventory, and pricing in an organized way.

### Data Modeling

The system uses the **Excel Data Model** to connect multiple tables through relationships. Instead of keeping data in separate spreadsheets, the tables were structured so that they behave more like a relational database.

The main tables in the system include:

- Products Table  
- Sales Table  
- Purchases Table  
- Inventory Table  

These tables are connected using **Product ID**, which allows Excel to combine information across different tables when generating reports and dashboard insights.

![Data Model](images/data_model.png)

---

### Analytical Thinking

The system converts raw operational data into meaningful business metrics.

Using **DAX measures**, several key indicators were created to help evaluate the performance of the business, including:

- Total Revenue  
- Cost of Goods Sold  
- Total Profit  
- Profit Margin  
- Total Orders  
- Current Stock Quantity  
- Cost of Goods Available  
- Stock Risk Count  

These metrics power the KPI cards and visual insights shown on the dashboard.

---

### Dashboard Design

A dashboard was created to present business performance in a format that is easy for a business owner to understand.

The dashboard provides insights into:

- revenue performance over time  
- most profitable products  
- sales channel performance  
- inventory health and stock risks  

Interactive filters allow users to analyze performance by **category, brand, and time period**.

---

### Automation with VBA

Several parts of the system were automated using **VBA macros** to reduce manual work and improve efficiency.

Automation was implemented for:

- submitting sales transactions
- submitting purchase transactions
- clearing entry forms
- deleting incorrect entries
- refreshing the dashboard and pivot tables

These automations ensure that the system operates smoothly without requiring manual data manipulation.

---

### How to Use the System

This system was designed so that the business owner interacts mainly with **entry forms and the dashboard**, while the calculations, relationships, and reporting happen automatically in the background. The steps below describe how the system can be used during normal daily operations.

---

#### 1. Register Products

Before recording any sales or purchases, all products must first be registered in the **Products Table**.  
This table acts as the master reference for every item sold in the store.

For each product, the following information should be entered:

- Category  
- Brand  
- Product Name  
- Size  
- Cost Price  
- Selling Price  

Each product automatically receives a **unique Product ID**, which is used throughout the system to connect the product with sales transactions, purchases, and inventory calculations.

Once products are registered here, they automatically become available in the dropdown lists used in the entry forms.

---

#### 2. Record Customer Sales

Whenever a customer purchases a product, the transaction is recorded using the **Sales Entry Form** rather than typing directly into the Sales table.

To record a sale:

1. Select the **Category** of the product  
2. Select the **Brand**  
3. Select the **Product Name**  
4. Enter the **Customer Name**  
5. Enter the **Quantity Purchased**  
6. Select the **Sales Channel**  
7. Click **Submit**

Once the Submit button is clicked, the system automatically records the transaction in the **Sales Table**.

Several fields are automatically generated or retrieved from the Products table, including:

- Product ID  
- Product Size  
- Product Price  
- Transaction Date  
- Sales ID  

This automation ensures that transactions are recorded consistently and reduces the risk of manual data entry errors.

Additional controls in the form include:

- **Clear Button** – resets the form so a new order can be recorded for the next customer  
- **Delete Button** – removes the most recent transaction if it was entered incorrectly  

---

#### 3. Record Inventory Purchases

When the business owner restocks products from suppliers or vendors, the transaction is recorded through the **Purchase Entry Form**.

To record a purchase:

1. Select the **Category**  
2. Select the **Brand**  
3. Select the **Product Name**  
4. Enter the **Vendor Name**  
5. Enter the **Quantity Purchased**  
6. Click **Submit**

Once submitted, the system records the purchase in the **Purchases Table**, increasing the available inventory for that product.

This allows the system to track how much stock is entering the business over time.

---

#### 4. Update Product Prices

If supplier costs change, the business owner can update product prices using the **Product Price Update Module**.

To update prices:

1. Select the product from the update form  
2. Enter the **new cost price** or **new selling price**  
3. Click the **Update Button**

The system automatically updates the corresponding values in the **Products Table**, ensuring that future transactions use the updated prices.

This prevents incorrect profit calculations when supplier prices change.

---

#### 5. Monitor Business Performance

The **Dashboard** provides a visual overview of business activity and performance.

The dashboard allows the business owner to monitor:

- total revenue generated
- cost of goods sold
- total profit and profit margin
- number of customer orders
- most profitable products
- sales performance by channel
- inventory value and stock availability
- products that are low in stock or out of stock

These insights allow the business owner to quickly understand the financial and operational health of the business.

---

#### 6. Apply Filters for Deeper Analysis

The dashboard includes interactive filters that allow the user to analyze performance from different perspectives.

Available filters include:

- **Category Filter**
- **Brand Filter**
- **Date Timeline**

For example:

If the business owner selects **Body Spray** in the Category filter, the dashboard immediately updates to show only the performance of body spray products.

The user can then filter further by **Brand** to analyze the performance of a specific brand within that category.

This filtering capability helps answer questions such as:

- Which product category generates the most revenue?
- Which brand is performing best?
- How have sales changed over time?

---

#### 7. Refresh the Dashboard

After recording new sales or purchase transactions, the business owner should click the **Refresh Button** on the dashboard.

This triggers a VBA macro that refreshes the entire data model and updates:

- pivot tables
- KPI cards
- charts
- inventory calculations

This ensures that the dashboard always reflects the **most recent business activity**.

---

By following this workflow, the system allows the business owner to efficiently record daily operations while automatically generating insights about sales performance, profitability, and inventory status.



##  Conclusion

This project presents a complete **Excel-based inventory and sales management system** designed for a perfume retail business. The goal of the system was to replace manual record keeping and guesswork with a structured process that allows the business owner to track sales, manage inventory, monitor profitability, and make better operational decisions.

The system was built using **Microsoft Excel, Power Pivot (Data Model), DAX calculations, and VBA automation** to simulate a small business intelligence workflow inside a spreadsheet environment.

At the core of the system is a structured **data model** connecting four main tables:

- Products Table – stores master information about each product including category, brand, size, cost price, and selling price  
- Sales Table – records every customer purchase transaction  
- Purchases Table – records restocking transactions from suppliers  
- Inventory Table – calculates current stock levels using purchase and sales activity  

These tables are connected using **Product ID**, allowing Excel to combine data across multiple tables and behave more like a relational database rather than separate spreadsheets.

To simplify daily operations for the business owner, the system includes **automated entry forms**:

- A **Sales Entry Form** used to record customer orders without typing directly into the data tables  
- A **Purchase Entry Form** used to record inventory restocking from vendors  
- A **Product Price Update Module** that allows the business owner to update cost and selling prices when supplier costs change  

The forms automatically retrieve product information such as **product ID, size, and pricing** from the Products table and record transactions directly into the corresponding tables.

Several operational tasks were automated using **VBA macros**, including:

- submitting sales and purchase transactions  
- clearing entry forms between orders  
- deleting incorrect transaction entries  
- refreshing the dashboard and pivot model  

These automations reduce manual work and help maintain consistent and reliable data records.

Using the **Excel Data Model and DAX measures**, the system calculates key business metrics such as:

- Total Revenue  
- Cost of Goods Sold (COGS)  
- Total Profit  
- Profit Margin  
- Total Orders  
- Current Stock Quantity  
- Cost of Goods Available (inventory value)  
- Stock Risk indicators for low or out-of-stock products  

These metrics power an **interactive dashboard** that visualizes business performance through KPI cards, revenue trend charts, product profitability analysis, sales channel performance, and inventory monitoring.

The dashboard also includes **interactive filters for category, brand, and date**, allowing the business owner to drill down into specific product segments and analyze performance from multiple perspectives. A **VBA-powered refresh button** ensures that the dashboard always reflects the most recent transactions recorded in the system.

Overall, this project demonstrates how structured data modeling, automation, and analytical reporting can be implemented using Excel to create a practical solution for managing real business operations. The system allows the business owner to move from manual tracking toward a **data-driven approach to monitoring sales performance, managing inventory levels, and making informed business decisions**.



##  Author & Contact

**Olivia Michael**

Data Analyst focused on using data to uncover business risks, performance patterns, and decision-making insights.

LinkedIn: [Olivia Anetoh](https://www.linkedin.com/in/olivia-anetoh-955b94328/)  
GitHub: [Olivia-Michael](https://github.com/Olivia-Michael)  
Email: anetohchinecherem@gmail.com


