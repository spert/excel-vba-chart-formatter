# Excel VBA Chart Formatter

The VBA code in this repository helps make Excel charts visually appealing—ideal for publications and presentations.

---

## Overview

Preparing Excel charts for regular publications or presentations can be tedious and time-consuming. Spreadsheets often contain raw data and calculations that are irrelevant to colleagues or editors. This VBA tool automates the routine editorial process of chart formatting.

The program:

- Copies the data from any selected Excel chart to a new worksheet
- Applies consistent formatting automatically
- Allows the creation of one or two chart copies (e.g., for multilingual use)
- Offers two size options: small (e.g., for print) or large (presentation-ready)
- Supports different title placements: inside a text box, as a standard chart title, or with no title

> ⚠️ **Note:** This tool is designed to prettify **simple chart types** (e.g., _line_ and _column_ charts). It may not work correctly with **complex chart types** such as _pie charts_, _radar charts_, or _scatter plots_.

---

## Usage

1. **Activate the chart**  
   Click once on the surface of the Excel chart you want to prettify.

2. **Click 'Format chart'**  
   Go to the **Add-ins** menu bar in Excel and click the **Format chart** button.  
   ![Chart Formatter Button](images/img1.png)

3. **Accept trust prompts**  
   Accept any messages related to macro or trust settings.

4. **Customize formatting**  
   A user form will open. Choose your preferred formatting options.  
   ![User Form](images/img2.png)

5. **Click 'Execute'**  
   Press the **Execute** button to start formatting.

6. **View the output**  
   A new worksheet will be created, containing:
   - The formatted chart (or two copies if selected)
   - The cleaned and copied chart data  
   ![Formatted Output](images/img4.png)

---

## Installation

1. Download the Excel add-in file (`ChartFormatter.xlam`) from this repository.
2. Open any Excel workbook.
3. Double-click the `ChartFormatter.xlam` file located in the `Install` folder to load the add-in.

---

## Technical Info

- The VBA code is compiled into a `.xlam` Excel Add-in file.
- Object-oriented programming principles are applied using VBA **classes**, **types**, and **collections**.
- Tested on **Excel 2016 Professional Plus**.
- Licensed under **GPLv3**. Please review the license terms before modifying or distributing the source code.

---



