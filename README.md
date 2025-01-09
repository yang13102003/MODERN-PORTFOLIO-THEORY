# Modern Portfolio Theory (MPT) - Excel with Macros

This repository provides an Excel-based implementation of **Modern Portfolio Theory (MPT)** using VBA macros. The tool is designed to help investors construct an optimized portfolio of 5 stocks, balancing risk and return efficiently.

## Features

- **Portfolio Optimization**: Uses MPT principles to calculate optimal stock allocation.
- **Efficient Frontier**: Visualizes the trade-off between risk and return.
- **Customizable Inputs**: Users can input their own stock data, constraints, and risk-free rates.
- **Automated Calculations**: Automates risk, return, covariance, and portfolio weight calculations using VBA.

## Requirements

- Microsoft Excel (2016 or later recommended)
- Enable macros in Excel settings

## Installation

1. **Download the Repository**:
   - Clone the repository using Git:
     ```bash
     git clone https://github.com/yourusername/mpt-excel-macros.git
     ```
   - Or download the ZIP file and extract it.

2. **Open the Excel File**:
   - Open the file `MPT_Portfolio_Optimization.xlsm` in Microsoft Excel.

3. **Enable Macros**:
   - Follow the prompts to enable macros in Excel.

## Usage

1. **Input Stock Data**:
   - Enter the historical price data for 5 stocks in the `Stock Data` worksheet.
   - Ensure the format includes dates in one column and price data in adjacent columns.

2. **Set Parameters**:
   - In the `Parameters` section, input:
     - Risk-free rate
     - Constraints (e.g., minimum or maximum weight for a stock)

3. **Run the Optimization**:
   - Click the "Run Optimization" button on the Excel sheet.
   - The macro will:
     - Calculate optimal portfolio weights.
     - Display the expected return, risk, and Sharpe ratio.

4. **View the Efficient Frontier**:
   - The efficient frontier chart is updated automatically to reflect the optimized portfolio.
