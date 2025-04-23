# 📊 Sales Dataset Analysis - Advanced Excel Project

## 📁 Project Overview

This project demonstrates **advanced Microsoft Excel techniques** to analyze a superstore sales dataset. The primary goal is to clean, process, and analyze sales data to uncover **actionable business insights** and provide **recommendations** for improving sales performance and optimizing discount strategies.

This project is aligned with the **"Advanced Excel Reinforcement Project"** and utilizes a dataset with **9,994 sales records**.

---

## 🧾 Dataset Information

- **File Name:** `Superstores sales dataset.xlsx`  
- **Total Records:** 9,994  
- **Columns:** 21  
  - Includes: `Order ID`, `Order Date`, `Ship Mode`, `Customer ID`, `Product Name`, `Sales`, `Quantity`, `Discount`, `Profit`, etc.  
- **Source:** Provided for educational purposes (not included in this repository due to potential copyright/privacy restrictions)

---

## 📌 Key Metrics

- **Sales:** Total revenue per order  
- **Adjusted Sales:** `Sales * (1 - Discount)`  
- **Profit:** Net profit per order  
- **Discount:** Percentage discount applied  

---

## 🎯 Project Objectives

- Clean and format the dataset
- Analyze sales performance by ship mode, product category, and time
- Evaluate discount impact on profitability
- Create PivotTables, charts, and an interactive dashboard
- Perform what-if analysis and automate tasks with macros
- Build a Power Pivot data model
- Derive insights and provide actionable recommendations

---

## 🛠️ Steps Performed

### 1. Data Cleaning and Formatting
- Removed duplicates (none found)
- Formatted dates (`Order Date`, `Ship Date`) and currency (`Sales`, `Discount`, `Profit`)
- Applied data validation rules:
  - Sales ≥ 0
  - Discount between 0 and 1
- Used `IF`/`IFERROR` formulas for error flagging

### 2. Sales Performance Analysis
- Calculated **Adjusted Sales**: `Sales * (1 - Discount)`
- Created PivotTables for:
  - **Ship Mode** (e.g., *Standard Class*: ~$1.2M Adjusted Sales)
  - **Product Categories** (e.g., *Technology*: ~$800K Adjusted Sales)
  - **Time Trends** (e.g., *November peak*: ~$200K)

### 3. Discount Impact Analysis
- Flagged high discounts (> 0.5) linked to losses, especially in **Furniture**
- Simulated reduced discount strategies to improve profitability

### 4. Visualization and Dashboard
- Built PivotCharts:
  - Bar charts for categories
  - Line charts for trends
- Designed a dashboard with KPIs and slicers

### 5. What-If Analysis
- **Goal Seek:** To achieve $500,000 profit target
- **Scenario Manager:** To test sales increase and discount reduction strategies

### 6. Automation
- Recorded **macros** for formatting and refreshing PivotTables
- Assigned macros to buttons for ease of use

### 7. Power Pivot and Data Modeling
- Created **Product** and **Customer** dimension tables
- Built relationships via `Product ID` and `Customer ID`
- Added a calculated column:  
  `Profit Margin = Profit / Sales`

---

## 📈 Key Insights

- **Ship Mode:**
  - *Standard Class* leads with ~$1.2M Adjusted Sales
  - *Same Day* performs lowest at ~$150K

- **Product Categories:**
  - *Best:* Technology (~$800K, high profit margins)
  - *Worst:* Furniture (~$600K, frequent losses due to discounts)
  - *Office Supplies:* High volume (~25K units), lower value per order

- **Discounts:**
  - Discounts > 0.5 reduce Adjusted Sales and cause losses  
    _E.g., Row 15: 0.8 discount → -$123.86 profit_

- **Seasonality:**
  - November shows peak sales (~$200K), likely due to holiday demand

---

## 📢 Recommendations

- **Cap Furniture Discounts:** Max 0.3 to reduce losses
- **Promote Technology Products:** Launch campaigns for high-margin items (e.g., Canon Copiers)
- **Loyalty Programs for Standard Class:** Offer 5% discount to retain customers
- **November Promotions:** Increase stock and run offers for Technology & Office Supplies
- **Optimize Discount Strategy:**  
  - Furniture: Max 0.3  
  - Office Supplies: Max 0.2  
  - Technology: Max 0.1  

---

## 📂 Repository Contents

- **Excel File:** `Sales_Dataset_Analysis.xlsx`  
  Contains:
  - Cleaned data
  - PivotTables
  - Charts and dashboards
  - Macros
  - Power Pivot data model

- **Sheets:**
  - `Sheet1`: Raw data + Adjusted Sales + Error Checks
  - `Summary`: KPIs and key insights
  - `Ship Mode Analysis`: PivotTable & Chart
  - `Category Analysis`: PivotTable & Chart
  - `Discount Analysis`: Loss scenarios
  - `Time Analysis`: Seasonal sales trends
  - `Marketing Plan`: Suggested promotions
  - `Dashboard`: Interactive with KPIs & slicers
  - `Products/Customers`: Power Pivot dimension tables

- **Screenshots (Folder: `/screenshots/`):**
  - `dashboard.png`: Dashboard view
  - `category_analysis.png`: Product category analysis
  - `time_trends.png`: Monthly sales trends

---

## 🧪 How to Use

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/your-username/sales-dataset-analysis.git

2. **Open the Excel File:**
   - Open `Sales_Dataset_Analysis.xlsx`
   - Ensure **Microsoft Excel** is installed (Microsoft 365 recommended)
   - Enable **Macros** if prompted

3. **Explore the Analysis:**
   - **Sheet1:** Review raw data, formulas, and Adjusted Sales
   - **Summary:** Check KPIs and insights
   - **Dashboard:** Use slicers to filter by category or ship mode
   - **Macros:** Use buttons to reformat or refresh analysis

4. **Replicate with Your Own Dataset:**
   - Ensure columns like: `Order ID`, `Order Date`, `Ship Mode`, `Sales`, `Profit`, `Discount`
   - Follow the steps and structure in this workbook

---

## 💻 Requirements

- **Software:** Microsoft Excel (with Power Pivot support)
- **Skills Needed:**  
  - Intermediate Excel knowledge (formulas, PivotTables, charts)  
  - Familiarity with macros and Power Pivot is helpful  
- **Dataset Format:** CSV or Excel with Superstore-like structure

---

## 🔮 Future Improvements

- Integrate **returns data** to refine profitability analysis
- Automate discount strategies using **VBA**
- Export Excel dashboard to **Power BI** for richer visuals

---

## 🙏 Acknowledgments

- Dataset provided for the **Advanced Excel Reinforcement Project**
- Inspired by real-world sales analytics use cases

---

## 📬 Contact

For questions, feedback, or contributions, feel free to:
- Raise an issue via GitHub  
- Email: rahmanmohammed620@gmail.com 
