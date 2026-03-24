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
