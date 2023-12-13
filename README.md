# 1. Mortgage Schedule Calculator Project
Important notes: This .xlsx project file works in excel desktop 2021 or later, excel 365, & google sheets. There are links to the project hosted in my OneDrive and google drive where you can try out a live version of my project without needing to download the .xlsx file. You can also download the .xlsx file and use it in your desktop version of excel. You also have the option of uploading the file to your OneDrive to use in excel on the web or upload the file to your google drive to use with google sheets. It is not necessary to install anything in order to understand this project because this README document will explain everything.
# 2. Personal Case
## a. Scenario Overview
My sister is in the process of buying a home and needs help understanding how home loans and full amortization works before she decides on her purchase. Since she knows that I'm a data analyst, she has asked me for my help with building her a tool that will help her make the best possible decision when deciding on a fixed mortgage for her house.

## b. Project Objective
The objective of this project is to build a tool that visually allows my sister to visualize and test the impact of different mortgage conditions and additional principal pre-payments on the mortgage payment schedule. It will also include a simple budgeting tool to help her make sure that her mortgage payments will be able to fit into her budget. The project will be built in an excel spreadsheet since my sister is familiar with the tool and since excel is a great tool for data analysis.

## c. Project Requirements
1. Use excel as the primary tool and also work in google sheets.
1. Use data validation to control what cells the users can edit and what data those editable cells can contain.
1. Create a table that displays the important mortgage information. It should allow the users to enter the following parameters that are used for the monthly mortgage payment calculation & the total monthly payments while PMI is inactive & active calculations (Mortgage Information):
    * House Price
    * Down Payment %
    * Number of years for the mortgage
    * Annual Interest Rate (Fixed)
    * Additional monthly costs (taxes, utilities, maintenance, etc.)
    * PMI Cutoff %
    * PMI % Annual
The section will 
1. Create a loan amortization payment schedule that displays the payment period number, date, starting principal balance, interest owed, principal owed, additional pre-payments made, principal remaining, additional costs owed, PMI owed, and total monthly payment for each payment period.
1. For each payment period, allow the users to enter in additional principal pre-payments to be included in the schedule's calculations for each payment period.
1. For each payment period, include any additional costs owed (property taxes/maintenance) & PMI along with the principal to allow for a total monthly payment calculation.
1. Allow the users to ender the start date for the loan and display the payment date for each monthly payment period.
1. Use dynamic excel formulas & conditional formatting to automatically change the size the payment schedule automatically based on the length of the mortgage.
1. Create a table that displays the number of payments and how much money you will save if you make additional principal pre-payments (Savings Metrics).
1. Create a table that allows the user to enter their household net income and see their Mortgage Payment as a % of Budget, Mortgage + Pre-Payment as a % of Budget, & Mortgage + Pre-Payment + Additional Costs + PMI as a % of Budget.
# 3. Project Steps
All of these steps will be displayed in a hidden tab called FormulaText Hidden. All of the highlighter yellow cells and orange principal pre payment cells will display values from the Adjustable Mortgage Schedule sheet. Every other cell with an excel formula will use the formulatext() function and refer to the matching cell values on the Adjustable Mortgage Schedule sheet. This will allow me to display all the formula logic in screenshots embedded in this readme document without having to write out the code for each excel formula. You can also just download the excel file on your own computer and look at all the formulas yourself to see how it all works. 

To use this file correctly and make sure that all the calculations work, the user simply needs to enter in the the required arguements with valid values into the highlighter yellow cells. The orange colored cells in the additional principal pre-payment column are optional arguemts for other calculations on the sheet. This means that the user can leave them blank and everything will work fine. Also, the number of orange cells available will vary based on the length of the payment schedule.



## a. Create the Mortgage Information Table
<details>
<summary>Mortgage Information Table with Formulas</summary>

![This table displays the important mortgage information.](images/FormulaText/Mortgage%20Information%20Table.jpg)

</details>

<details>
<summary>Mortgage Information Table with Output</summary>

![This table displays the important mortgage information.](images/Output/Mortgage%20Information%20Table.jpg)

</details>

This table displays the important mortgage information. It allows the users to enter the following parameters that are used for the monthly mortgage payment calculation & the total monthly payments while PMI is inactive & active calculations (Mortgage Information):

* House Price
* Down Payment %
* Number of years for the mortgage
* Annual Interest Rate (Fixed)
* Additional monthly costs (taxes, utilities, maintenance, etc.)
* PMI Cutoff %
* PMI % Annual

Notice how all the required arguements are colored highlighter yellow.

## b. Create the Payment Schedule Dynamic Table

This table contains a loan amortization payment schedule that displays the payment period number, date, starting principal balance, interest owed, principal owed, additional pre-payments made, principal remaining, additional costs owed, PMI owed, and total monthly payment for each payment period. For each payment period, the users can optionally enter in additional principal pre-payments to be included in the schedule's calculations for each payment period. The additional prepayment cells are optional arguemts for other calculations in the sheet so they are colored orange.

Since each excel formula in this table is large, every column will get it's own image so we can see the logic of how it works. In most of the formulas in the payment schedule table, we use the LET() function with variables to make the logic of the formula easier to understand. Most of the formulas with LET() have these 3 variables in every cell. They are IsPeriodBlankCondition, IsPrincipalRemainingBlankCondition, & IsPrincipalZeroCondition. These 3 variables contain conditions that allow the calculations in the payment schedule table to dynamically decide if the cell should output an empty string or if it needs to perform a calculation for the payment period.

Each of these calculations begin on the 2nd row of the Payment Schedule table with the exception of the month column (3rd row) and the principal remaining column (1st row). In the pictures of each column below, we will show the 1st formula for each column using formulatext(). These formulas used mixed referencing to allow us to only write 1 formula for each column and autofill to the bottom of the table. 

<details>
<summary>Period Column with Formulas</summary>

![This column displays the payment period formula.](images/FormulaText/Payment%20Period%20Column.jpg)

The formula in this column is an array formula that adjusts in size based on the number of years that the mortgage is. This means that this column varys in size and that only cell B26 contains a formula in this column. This becuase the formula in cell B26 spills down column B. Also, every other column in the payment schedule table except for the additional principal pre-payment column 1st checks to see if the payment period column is empty becuase if so, then the result of the formula will be an empty string.

</details>

<details>
<summary>Month Column with Formulas</summary>

![This column displays the period month formula.](images/FormulaText/Month%20Column.jpg)

This column displays the period month formula.

</details>

<details>
<summary>Starting Principal Balance Column Formula</summary>

![This column displays the starting principal balance formula.](images/FormulaText/Starting%20Principal%20Balance%20Column.jpg)

This column displays the starting principal balance formula.

</details>

<details>
<summary>Interest Owed Column Formula</summary>

![This column displays the interest owed formula.](images/FormulaText/Interest%20Owed%20Column.jpg)

This column displays the interest owed formula.

</details>

<details>
<summary>Principal Owed Column Formula</summary>

![This column displays the principal owed formula.](images/FormulaText/Principal%20Owed%20Column.jpg)

This column displays the principal owed formula.

</details>

<details>
<summary>Additional Principal Pre-Payment Column Formula</summary>

![This column displays the additional principal pre-payment formula.](images/FormulaText/Additional%20Principal%20Pre-Payment%20Column.jpg)

This column displays the additional principal pre-payment formula.

</details>

<details>
<summary>Principal Remaining Formula</summary>

![This column displays the principal remaining formula.](images/FormulaText/Principal%20Remaining%20Column.jpg)

This column displays the principal remaining formula.

</details>

<details>
<summary>Additional Costs Owed Column Formula</summary>

![This column displays the additional costs owed formula.](images/FormulaText/Additional%20Costs%20Owed%20Column.jpg)

This column displays the additional costs owed formula.


</details>

<details>
<summary>PMI Owed Column Formula</summary>

![This column displays the PMI owed formula.](images/FormulaText/PMI%20Owed%20Column.jpg)

This column displays the PMI owed formula.

</details>

<details>
<summary>Total Monthly Payment Column Formula</summary>

![This column displays the total monthly payment formula.](images/FormulaText/Total%20Monthly%20Payment%20Column.jpg)

</details>

<details>
<summary>Payment Schedule Table</summary>

![This table displays the output of each of the columns in the payment schedule table.](images/Output/Payment%20Schedule%20Table.jpg)

This table displays the output of each of the columns in the payment schedule table.

</details>


## c. Create the Savings Metrics Table and Adjustable Mortgage Hidden Sheet

<details>
<summary>Savings Metrics Table with Formulas</summary>

![This table displays the important savings metrics information.](images/FormulaText/Savings%20Metrics%20Table.jpg)

</details>

<details>
<summary>Savings Metrics Table with Output</summary>

![This table displays the important savings metrics information.](images/Output/Savings%20Metrics%20Table.jpg)

This table displays the number of payments and how much money you will save if you make additional principal pre-payments (Savings Metrics).

</details>


<details>
<summary>Adjustable Mortgage Hidden Sheet with Output</summary>

![This hidden sheet displays a mirror of all the Mortgage Information Table & Payment Schedule Table calculations on the Adjustable Mortgage Payment Schedule sheet without the Additional Principal Pre-Payment column. This allows us see the savings impact when making additional principal pre-payments in some of the the Savings Metrics Table's calculations.](images/Output/Adjustable%20Mortgage%20Hidden%20Sheet.jpg)

This hidden sheet displays a mirror of all the Mortgage Information Table & Payment Schedule Table calculations on the Adjustable Mortgage Payment Schedule sheet without the Additional Principal Pre-Payment column. This allows us see the savings impact when making additional principal pre-payments in some of the the Savings Metrics Table's calculations.

</details>


## d. Create the Impact of House Payments on Budget Table

<details>
<summary>Impact of House Payments on Budget Table with Formulas</summary>

![This table displays the important House Payment Budget information formulas.](images/FormulaText/Impact%20of%20House%20Payments%20on%20Budget%20Table.jpg)

</details>

<details>
<summary>Impact of House Payments on Budget Table with Output</summary>

![This table displays the important House Payment Budget information metrics.](images/Output/Impact%20of%20House%20Payments%20on%20Budget%20Table.jpg)

</details>

This table allows the user to enter their household net income and see their Mortgage Payment as a % of Budget, Mortgage + Pre-Payment as a % of Budget, & Mortgage + Pre-Payment + Additional Costs + PMI as a % of Budget.

## e. Create Data Validation

For the data validation I did 3 things:
1. Use the allow edit ranges feature to limit which cells users can make edits to. These are all the yellow cells and cells that cen be orange (G26:G625). 
1. Create custom data validation for each of the editable cells with custom input messages and error alerts.
1. Protect all 3 worksheets (including the 2 hidden ones).


# 4. Results

# 5 Technical Concepts Used
* Excel - Dynamic array formulas, spilled ranges, dynamic array functions, date & time functions, information functions, financial functions, statistical functions, logical functions, exception handling, conditional formatting, freeze panes, formula precedent and dependent values, relative referencing, fixed referencing, mixed referencing, Let() variables, autofill, formulatext(), implicit boolean type coercion, hidden sheets, data validation, allow edit ranges, & worksheet protection.

* Finance - Mortgage, amortization, PMI(Private Mortgage Insurance), additional costs, fixed interest rate, principal, payment schedule, net income, budget, savings, starting balance, & payment periods.

# 6. How to run the Report on your Computer

# 7. Links to Code