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
