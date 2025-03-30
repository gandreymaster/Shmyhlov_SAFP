from read_xlsx import read_xlsx
import calc_values
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

firms = read_xlsx("Firms_Chestionar.xlsx")
print(len(firms))

# Бінарна класіфікація
def assess_company_status(data, turnover_growth, field_of_activity):
    current_ratio = calc_values.current_ratio(data["circulant_assets"], data["liabilities"])
    profit_margin = calc_values.profit_margin(data["profit_net"], data["turnover"])
    roa = calc_values.return_on_assets(data["profit_net"],
                                     calc_values.total_assets(data["fixed_assets"], data["circulant_assets"]))
    roe = calc_values.return_on_equity(data["profit_net"], data["capitals_and_reserves"])
    debt_to_equity = calc_values.debt_to_equity_ratio(data["liabilities"], data["capitals_and_reserves"])

    if pd.isna([current_ratio, profit_margin, roa, roe, debt_to_equity, turnover_growth]).any():
        return "Unknown", {"Current Ratio": "N/A", "Profit Margin": "N/A", "ROA": "N/A",
                          "ROE": "N/A", "Debt-to-Equity": "N/A"}

    metrics_assessment = {}
    main_category = list(field_of_activity.keys())[0]

    # Current Ratio
    optimal_cr_ranges = {
        "Transport and Logistics": (1.0, 1.5),
        "Construction": (1.2, 2.0),
        "Trade": (1.5, 2.5),
        "Services": (1.0, 1.5),
        "Medical and Healthcare": (1.2, 2.0),
        "IT and Technology": (1.0, 1.5),
        "Real Estate": (1.5, 2.5),
        "Automotive": (1.2, 1.8),
        "Uncategorized": (1.0, 2.0)
    }
    cr_min, cr_max = optimal_cr_ranges.get(main_category, optimal_cr_ranges["Uncategorized"])
    if current_ratio > cr_max:
        metrics_assessment["Current Ratio"] = "Excessive"
        cr_score = 1
    elif cr_min <= current_ratio <= cr_max:
        metrics_assessment["Current Ratio"] = "Optimal"
        cr_score = 2
    else:
        metrics_assessment["Current Ratio"] = "Insufficient"
        cr_score = 0

    # Profit Margin (%)
    optimal_pm_ranges = {
        "Transport and Logistics": (5, 10),
        "Construction": (5, 15),
        "Trade": (2, 10),
        "Services": (15, 30),
        "Medical and Healthcare": (10, 20),
        "IT and Technology": (20, 40),
        "Real Estate": (10, 25),
        "Automotive": (5, 15),
        "Uncategorized": (5, 20)
    }
    pm_min, pm_max = optimal_pm_ranges.get(main_category, optimal_pm_ranges["Uncategorized"])
    if profit_margin > pm_max:
        metrics_assessment["Profit Margin"] = "Exceptional"
        pm_score = 3
    elif pm_min <= profit_margin <= pm_max:
        metrics_assessment["Profit Margin"] = "Optimal"
        pm_score = 2
    elif profit_margin > 0:
        metrics_assessment["Profit Margin"] = "Suboptimal"
        pm_score = 1
    else:
        metrics_assessment["Profit Margin"] = "Loss-Making"
        pm_score = 0

    # ROA (%)
    optimal_roa_ranges = {
        "Transport and Logistics": (3, 8),
        "Construction": (2, 6),
        "Trade": (5, 10),
        "Services": (10, 20),
        "Medical and Healthcare": (5, 15),
        "IT and Technology": (15, 30),
        "Real Estate": (2, 8),
        "Automotive": (3, 10),
        "Uncategorized": (5, 15)
    }
    roa_min, roa_max = optimal_roa_ranges.get(main_category, optimal_roa_ranges["Uncategorized"])
    if roa > roa_max:
        metrics_assessment["ROA"] = "Exceptional"
        roa_score = 3
    elif roa_min <= roa <= roa_max:
        metrics_assessment["ROA"] = "Optimal"
        roa_score = 2
    elif roa > 0:
        metrics_assessment["ROA"] = "Suboptimal"
        roa_score = 1
    else:
        metrics_assessment["ROA"] = "Unprofitable"
        roa_score = 0

    # ROE (%)
    optimal_roe_ranges = {
        "Transport and Logistics": (8, 15),
        "Construction": (10, 20),
        "Trade": (10, 18),
        "Services": (15, 25),
        "Medical and Healthcare": (12, 20),
        "IT and Technology": (20, 35),
        "Real Estate": (8, 18),
        "Automotive": (10, 20),
        "Uncategorized": (10, 20)
    }
    roe_min, roe_max = optimal_roe_ranges.get(main_category, optimal_roe_ranges["Uncategorized"])
    if roe > roe_max:
        metrics_assessment["ROE"] = "Exceptional"
        roe_score = 2
    elif roe_min <= roe <= roe_max:
        metrics_assessment["ROE"] = "Optimal"
        roe_score = 1
    else:
        metrics_assessment["ROE"] = "Suboptimal"
        roe_score = 0

    # Debt-to-Equity
    optimal_dte_ranges = {
        "Transport and Logistics": (1.0, 2.0),
        "Construction": (1.0, 1.8),
        "Trade": (0.5, 1.5),
        "Services": (0.3, 1.0),
        "Medical and Healthcare": (0.5, 1.5),
        "IT and Technology": (0.2, 1.0),
        "Real Estate": (1.5, 3.0),
        "Automotive": (1.0, 2.0),
        "Uncategorized": (0.5, 1.5)
    }
    dte_min, dte_max = optimal_dte_ranges.get(main_category, optimal_dte_ranges["Uncategorized"])
    if debt_to_equity < dte_min:
        metrics_assessment["Debt-to-Equity"] = "Very Low"
        dte_score = 1
    elif dte_min <= debt_to_equity <= dte_max:
        metrics_assessment["Debt-to-Equity"] = "Optimal"
        dte_score = 2
    else:
        metrics_assessment["Debt-to-Equity"] = "High"
        dte_score = 0

    total_score = cr_score + pm_score + roa_score + roe_score + dte_score

    if total_score >= 9:
        overall_status = "Good"
    elif total_score >= 5:
        overall_status = "Moderate"
    else:
        overall_status = "Bad"

    return overall_status, metrics_assessment

# Z-score Альтмана
def calculate_z_score(data, ismanufacturing):
    total_assets = calc_values.total_assets(data["fixed_assets"], data["circulant_assets"])

    if total_assets == 0 or np.isnan(total_assets):
        return np.nan, "Unable to Calculate"

    X1 = (data["circulant_assets"] - data["liabilities"]) / total_assets
    X2 = data["capitals_and_reserves"] / total_assets
    X3 = data["profit_net"] / total_assets
    X4 = data["capitals_and_reserves"] / data["liabilities"] if data["liabilities"] != 0 else np.inf
    X5 = data["turnover"] / total_assets

    if ismanufacturing:
        z_score = 1.2 * X1 + 1.4 * X2 + 3.3 * X3 + 0.6 * X4 + 1.0 * X5
        thresholds = {"Safe Zone": 2.99, "Grey Zone": 1.8}
    else:
        z_score = 6.56 * X1 + 3.26 * X2 + 6.72 * X3 + 1.05 * X4
        thresholds = {"Safe Zone": 2.6, "Grey Zone": 1.1}

    if pd.isna(z_score):
        return np.nan, "Unable to Calculate"
    elif z_score > thresholds["Safe Zone"]:
        return z_score, "Safe Zone"
    elif thresholds["Grey Zone"] <= z_score <= thresholds["Safe Zone"]:
        return z_score, "Grey Zone"
    else:
        return z_score, "Bankruptcy Zone"


results = []
for firm in firms:
    firm_name = firm["Name"]
    years = firm["Years"]

    for i, year_data in enumerate(years):
        data = {
            "turnover": year_data["Turnover"],
            "profit_net": year_data["Profit Net"],
            "liabilities": year_data["Liailities"],
            "fixed_assets": year_data["Fixed assets"],
            "circulant_assets": year_data["Circulant Assets"],
            "capitals_and_reserves": year_data["Capitals and reserves"],
            "average_number_of_employees": year_data["The average number of employees"]
        }

        turnover_growth = 0
        if i > 0:
            prev_turnover = years[i - 1]["Turnover"]
            turnover_growth = (data["turnover"] - prev_turnover) / prev_turnover if prev_turnover != 0 else 0

        overall_status, metrics_assessment = assess_company_status(data, turnover_growth, firm['Field of activity'])

        z_score, z_category = calculate_z_score(data, firm['Is manufacturing?'])

        current_ratio = round(calc_values.current_ratio(data["circulant_assets"], data["liabilities"]), 2)
        profit_margin = round(calc_values.profit_margin(data["profit_net"], data["turnover"]), 2)
        roa = round(calc_values.return_on_assets(data["profit_net"],
                                                 calc_values.total_assets(data["fixed_assets"],
                                                                          data["circulant_assets"])), 2)
        roe = round(calc_values.return_on_equity(data["profit_net"], data["capitals_and_reserves"]), 2)
        debt_to_equity = round(calc_values.debt_to_equity_ratio(data["liabilities"], data["capitals_and_reserves"]), 2)

        result = {
            "Company": firm_name,
            "Year": year_data["Year"],
            "Current Ratio": current_ratio,
            "Current Ratio Status": metrics_assessment["Current Ratio"],
            "Profit Margin (%)": profit_margin,
            "Profit Margin Status": metrics_assessment["Profit Margin"],
            "ROA (%)": roa,
            "ROA Status": metrics_assessment["ROA"],
            "ROE (%)": roe,
            "ROE Status": metrics_assessment["ROE"],
            "Debt-to-Equity": debt_to_equity,
            "Debt-to-Equity Status": metrics_assessment["Debt-to-Equity"],
            "Z-score": round(z_score, 2) if pd.notna(z_score) else z_score,
            "Z-category": z_category,
            "Overall Status": overall_status,
        }
        results.append(result)

results_df = pd.DataFrame(results)
print(results_df.head())
results_df.to_csv("firm_analysis_results.csv", index=False)

print("\nПорівняння оцінок:")
crosstab = pd.crosstab(results_df["Overall Status"], results_df["Z-category"])
print(crosstab)

plt.figure(figsize=(10, 6))
crosstab.plot(kind='bar', stacked=True, colormap='viridis')
plt.title("Overall Status vs Z-category Distribution")
plt.xlabel("Overall Status")
plt.ylabel("Number of Companies")
plt.xticks(rotation=0)
plt.legend(title="Z-category")
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show()
