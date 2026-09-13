# 10-Game Poisson BTTS Matrix & Staking Calculator

An automated sports analytics tool that uses the **Poisson Distribution** to calculate true probabilities for the Both Teams to Score (BTTS) football market. 

It generates a highly customized Excel workbook (`.xlsx`) formatted natively in **Nigerian Naira (₦)**, featuring a built-in Fractional Kelly Staking calculator and a Closing Line Value (CLV) performance tracker.

## 📊 Features Included
* **Poisson Shortcuts:** Calculates joint probabilities using mathematical shortcuts based on team expected goals ($\lambda$).
* **Fractional Kelly Staking:** Prevents bankroll ruin by dynamic sizing based on your exact value edge.
* **CLV Tracker:** Measures your execution sharpness against the closing bookmaker lines.

## 🚀 How to Run
1. Install dependencies:
   ```bash
   pip install openpyxl pandas
   ```
2. Run the generator script:
   ```bash
   python main.py
   ```
3. Open the newly generated `btts_poisson_clv_naira_calculator.xlsx` file in Microsoft Excel or Google Sheets.

