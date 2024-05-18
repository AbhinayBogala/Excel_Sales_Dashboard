# Excel_Sales_Dashboard
The goal of this project is to develop an interactive and comprehensive sales dashboard using Excel. This dashboard will offer a streamlined, visual approach to analyzing sales data, enabling stakeholders to make data-driven decisions efficiently.

![Screenshot 2024-05-18 191704](https://github.com/AbhinayBogala/Excel_Sales_Dashboard/assets/168276269/058ccadc-0bde-4fb7-9616-821e1166f91a)


## Objective of the Project
The objective of this project is to design and implement an interactive and comprehensive sales dashboard using Excel. 

### Key Points of the Sales Dashboard Project

1. **Data Cleaning**: Ensures the dataset is accurate, consistent, and free of duplicates, missing values, and inconsistencies, laying a strong foundation for reliable analysis.

2. **Pivot Table Utilization**: Leverages the power of pivot tables to summarize and analyze sales data across various dimensions such as customer, state, category, sub-category, product, sales, quantity, and profit.

3. **Interactive Dashboard Design**: Creates an interactive and user-friendly dashboard with visual elements like charts and slicers, enabling stakeholders to explore data dynamically and intuitively.

4. **Comprehensive Analysis**: Provides in-depth insights into sales performance, identifying top-performing products, high-revenue customers, and key geographic areas, facilitating targeted decision-making.

5. **Enhanced Data Visibility**: Offers a clear and organized view of sales metrics, making it easier for stakeholders to understand and interpret data at a glance.

6. **Support for Strategic Decisions**: Empowers stakeholders with actionable insights to drive strategic planning, resource allocation, and operational improvements, ultimately aiming to boost sales and profitability.

7. **Efficiency and Accuracy**: Streamlines the data analysis process, saving time and reducing errors, thus ensuring that the insights derived are both timely and accurate.



## Stages to Create a Sales Dashboard in Excel

Creating a sales dashboard in Excel involves several stages, from initial data preparation to final visualization. Below is a step-by-step guide:

#### 1. **Data Collection**
   - **Gather Data**: Compile the necessary sales data from various sources such as CRM systems, ERP systems, or databases.
   - **Consolidate Data**: Ensure all data is in a single, consistent format and location, typically in an Excel worksheet.

#### 2. **Data Cleaning**
   - **Remove Duplicates**: Identify and eliminate duplicate records.
   - **Handle Missing Values**: Address missing data by either filling in gaps or removing incomplete records.
   - **Standardize Data**: Ensure consistency in data formats (e.g., dates, currency, text).

#### 3. **Data Structuring**
   - **Organize Data**: Structure the data in a tabular format with clear headings for each column (e.g., Customer Name, State, Category, Sub-Category, Product Name, Sales, Quantity, Profit).
   - **Create Data Table**: Convert the dataset into an Excel table for better manageability and automatic updates (Insert > Table).

#### 4. **Pivot Table Creation**
   - **Insert Pivot Table**:
     - Select the data range and choose `Insert` > `PivotTable`.
     - Place the PivotTable in a new worksheet.
   - **Configure Pivot Table Fields**:
     - Drag `Customer Name`, `State`, `Category`, `Sub-Category`, and `Product Name` to the Rows area.
     - Use `State` or `Category` in the Columns area if needed.
     - Drag `Sales`, `Quantity`, and `Profit` to the Values area and set them to sum.
   - **Adjust Field Settings**: Format the value fields for currency and numbers as needed.

#### 5. **Visualization**
   - **Insert Charts**:
     - Create charts based on PivotTable data (e.g., bar charts, pie charts, line charts).
     - Go to `Insert` > `Chart` and select the desired chart type.
   - **Add Slicers**:
     - Insert slicers for interactive filtering (Analyze > Insert Slicer) and choose fields like `Category`, `State`, and `Customer Name`.
   - **Design Dashboard Layout**:
     - Arrange pivot tables and charts on a new worksheet titled `Dashboard`.
     - Ensure a clean, user-friendly layout with clear labels and headings.

#### 6. **Interactivity and Formatting**
   - **Link Slicers**: Connect slicers to the relevant pivot tables for dynamic data filtering.
   - **Apply Conditional Formatting**: Use conditional formatting to highlight key metrics (Home > Conditional Formatting).
   - **Consistent Formatting**: Apply consistent font styles, colors, and formats across the dashboard for a professional look.

#### 7. **Validation and Testing**
   - **Validate Data**: Check the accuracy and integrity of the data and calculations.
   - **Test Functionality**: Ensure all interactive elements, such as slicers and charts, function correctly.

#### 8. **Final Review and Distribution**
   - **Review Layout**: Make any necessary adjustments to the layout and design.
   - **Document Insights**: Summarize key findings and insights that the dashboard reveals.
   - **Share Dashboard**: Distribute the dashboard to stakeholders or upload it to a shared platform for access.

By following these stages, you can create a robust and interactive sales dashboard in Excel that provides valuable insights and supports informed decision-making.



## 1. **Data Collection**
What data is needed to achieve our objective?
We need data on the Sales Of an Company that includes their

- Order Date
- Years
- Months
- Customer Name
- State	
- Category
- Sub-Category
- Product Name
- Sales
- Quantity
- Profit

Where is the data coming from? 
The data is sourced from Kaggle (an Excel extract).


## 2. **Data Cleaning**


### Steps for Data Cleaning and Checking in Excel

#### 1. **Remove Duplicates**
   - **Select Data Range**: Highlight the data range you want to check for duplicates.
   - **Remove Duplicates**: Go to `Data` > `Remove Duplicates`. Select columns to consider for duplication and click `OK`.

#### 2. **Handle Missing Values**
   - **Identify Missing Data**: Use filters to identify blank cells. Go to `Data` > `Filter`, then click the dropdown arrow in each column header and select `(Blanks)`.
   - **Fill or Remove**: Decide whether to fill missing values with appropriate data (e.g., average, median) or to remove the rows with missing data. Use `Find & Select` > `Go To Special` > `Blanks` for easier identification.

#### 3. **Standardize Data Formats**
   - **Date Formats**: Ensure all date entries are in a consistent format. Select the column, right-click, choose `Format Cells`, and select the desired date format.


## 3. **Data Structuring**
   - **Organize Data**: Structure the data in a tabular format with clear headings for each column (e.g., Customer Name, State, Category, Sub-Category, Product Name, Sales, Quantity, Profit).
   - **Create Data Table**: Convert the dataset into an Excel table for better manageability and automatic updates (Insert > Table).


## 4. **Pivot Table Creation**
1. Select the data range in Excel.
2. Navigate to the `Insert` tab and click on `PivotTable`.
3. Choose the option to place the PivotTable in a new worksheet.
4. In the PivotTable Fields pane, drag fields like `Customer Name`, `State`, `Category`, `Sub-Category`, and `Product Name` to the Rows area.
5. For additional analysis, place `State` or `Category` in the Columns area.
6. Drag `Sales`, `Quantity`, and `Profit` to the Values area and set them to sum by clicking on the field settings.
7. Right-click on the value fields in the PivotTable, select `Value Field Settings`, and format them for currency or numbers.
8. Utilize filters and slicers to refine the data presentation if necessary.
9. Customize PivotTable layout and design using Excel's formatting options.
10. Update the PivotTable as needed by refreshing the data source or adjusting field selections.

## 5. **Visualization**


1. **Ensure Data Structure**: Organize your data in a table with columns for Customer Name, State, Category, Sub-Category, Product Name, Sales, Quantity, Profit, Date, and Month.


![Screenshot 2024-05-18 193919](https://github.com/AbhinayBogala/Excel_Sales_Dashboard/assets/168276269/8b2ab3a7-bf5b-4df4-bea7-05c3c0b7c60d)




2. **Create Pivot Table for Sales by Category (Funnel Chart)**:
   - Select the data range.
   - Insert a PivotTable and place it in a new worksheet.
   - Drag 'Sub-Category' to Rows and 'Sales' to Values, set to sum.
   - Sort the Sub-Categories by Sales.
   - Create a funnel chart from this data.
   - 

![Screenshot 2024-05-18 194811](https://github.com/AbhinayBogala/Excel_Sales_Dashboard/assets/168276269/019a6dfe-f5b7-4024-ad6a-1aa2e1d8229c)




3. **Create Pivot Table for Profit Gained Over Time (Line Chart)**:
   - Insert another PivotTable in a new worksheet.
   - Drag 'Year' to Rows and 'Profit' to Values, set to sum.
   - Create a line chart from this data.
   - 

![Screenshot 2024-05-18 201237](https://github.com/AbhinayBogala/Excel_Sales_Dashboard/assets/168276269/580ecbd0-73eb-4752-aa68-425e7f045dae)




4. **Create Monthly Sales (Area Chart)**:
   - Create a PivotTable with 'Month' and 'Year' slicers.
   - Drag 'Month' to Row and 'Sales' to Values, set to sum.
   - Insert an area chart based on this PivotTable.
   - 

![Screenshot 2024-05-18 201321](https://github.com/AbhinayBogala/Excel_Sales_Dashboard/assets/168276269/a9d1a6df-3009-41c5-895d-ea5d8828f004)




5. **Top 5 Customers (Pie Chart)**:
   - Use a PivotTable with 'Customer Name' and 'Profit'.
   - Sort customers by Profit and select the top 5.
   - Create a pie chart from this data.
   - 

![Screenshot 2024-05-18 201338](https://github.com/AbhinayBogala/Excel_Sales_Dashboard/assets/168276269/bffea596-ba34-4742-8b15-bc2518590e9c)




6. **Sales by State (Map)**:
   - Utilize a PivotTable with 'State' and 'Sales'.
   - Insert a map chart and add a slicer for 'Category'.
   - 

![Screenshot 2024-05-18 201355](https://github.com/AbhinayBogala/Excel_Sales_Dashboard/assets/168276269/814a0f78-f951-439a-9de4-c428ef743909)




7. **Customer Chart (Sidebar Chart)**:
   - Create a Table with 'Year' and 'Customer Name'.
   - Remove duplicates from the 'Customer Name' field.
   - Generate a sidebar chart based on this PivotTable.
   

![Screenshot 2024-05-18 201409](https://github.com/AbhinayBogala/Excel_Sales_Dashboard/assets/168276269/8384586d-05c2-4fcf-a0e0-878e245a5564)




## 6. **Interactivity and Formatting**

**Position and Format Charts**:
  - Create a Slicer For Category and Years.
  - By right clicking on the Name we find slicer.
  - Move each chart to its respective worksheet.
  - Add titles, axes labels, and legends as necessary.
  - Format charts for clarity and consistency.

**Link Slicers**
    - Link Slicer to the all the categories in the dashboard.
    


## 7. **Validation and Testing**

   - Ensure that all slicers and filters update the charts correctly.
   - Check for any discrepancies or errors in the data or charts.

     

## 8. **Final Review and Distribution**
   - **Review Layout**: Make any necessary adjustments to the layout and design.
   - **Document Insights**: Summarize key findings and insights that the dashboard reveals.
   - **Share Dashboard**: Distribute the dashboard to stakeholders or upload it to a shared platform for access.
