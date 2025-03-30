import pandas as pd
import numpy as np


def current_ratio(circulant_assets, liabilities):
    return circulant_assets / liabilities if liabilities and pd.notna(liabilities) else np.nan


def profit_margin(profit_net, turnover):
    return (profit_net / turnover) * 100 if turnover and pd.notna(turnover) else np.nan


def return_on_assets(profit_net, total_assets):
    return (profit_net / total_assets) * 100 if total_assets and pd.notna(total_assets) else np.nan


def return_on_equity(profit_net, capitals_and_reserves):
    return (profit_net / capitals_and_reserves) * 100 if capitals_and_reserves and pd.notna(capitals_and_reserves) else np.nan


def debt_to_equity_ratio(liabilities, capitals_and_reserves):
    return liabilities / capitals_and_reserves if capitals_and_reserves and pd.notna(capitals_and_reserves) else np.nan


def asset_turnover(turnover, total_assets):
    return turnover / total_assets if total_assets and pd.notna(total_assets) else np.nan


def fixed_asset_turnover(turnover, fixed_assets):
    return turnover / fixed_assets if fixed_assets and pd.notna(fixed_assets) else np.nan


def labor_productivity(turnover, average_number_of_employees):
    return turnover / average_number_of_employees if average_number_of_employees and pd.notna(average_number_of_employees) else np.nan


def total_assets(fixed_assets, circulant_assets):
    return fixed_assets + circulant_assets if pd.notna(fixed_assets) and pd.notna(circulant_assets) else np.nan
