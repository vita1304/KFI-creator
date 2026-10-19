import io
import openpyxl
import pandas as pd


def parse_financial_statements(file_bytes: bytes):
    """Зчитує Форму 1 та Форму 2 за кодами рядків."""
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes))

    f1_sheet = next(
        (s for s in excel_file.sheet_names if "Форма № 1" in s or "Баланс" in s),
        None,
    )
    f2_sheet = next(
        (s for s in excel_file.sheet_names if "Форма № 2" in s or "Звіт про" in s),
        None,
    )

    if not f1_sheet or not f2_sheet:
        raise ValueError(
            "У файлі обов'язково мають бути аркуші 'Форма № 1' та 'Форма № 2'."
        )

    f1_raw = pd.read_excel(excel_file, sheet_name=f1_sheet)
    f2_raw = pd.read_excel(excel_file, sheet_name=f2_sheet)

    def clean_form(df):
        header_idx = df[
            df.apply(lambda r: r.astype(str).str.contains("Код").any(), axis=1)
        ].index[0]
        cleaned = df.iloc[header_idx + 1 :].copy()
        cleaned.columns = df.iloc[header_idx]
        cleaned["Код"] = pd.to_numeric(cleaned["Код"], errors="coerce")
        cleaned = cleaned.dropna(subset=["Код"])
        cleaned["Код"] = cleaned["Код"].astype(int)
        val_cols = [
            c for c in cleaned.columns if c not in ["Код", cleaned.columns[0]]
        ]
        data = {}
        for _, row in cleaned.iterrows():
            data[row["Код"]] = {
                "prev": pd.to_numeric(row[val_cols[0]], errors="coerce") or 0.0,
                "curr": pd.to_numeric(row[val_cols[1]], errors="coerce") or 0.0,
            }
        return data

    return clean_form(f1_raw), clean_form(f2_raw)


def calculate_kfi_data(f1_data: dict, f2_data: dict, eur_rate: float = 48.44):
    """Розраховує всі статті 4-го аркуша KFI відповідно до структури."""
    g1 = lambda code, p="curr": f1_data.get(code, {}).get(p, 0.0)
    g2 = lambda code, p="curr": f2_data.get(code, {}).get(p, 0.0)

    # 1. Top block (EUR & Growth)
    fixed_assets_eur = g1(1010, "curr") / eur_rate / 1000.0
    fixed_assets_pct = (
        g1(1010, "curr") / g1(1300, "curr") if g1(1300, "curr") else 0.0
    )
    turnover_growth = (
        (g2(2000, "curr") / g2(2000, "prev") - 1.0) if g2(2000, "prev") else 0.0
    )
    inv_real_estate = (
        (g1(1016, "curr") - g1(1017, "curr")) / eur_rate / 1000.0
    )  # якщо є

    # 2. P&L BRIDGE (Mio. UAH)
    turnover = g2(2000, "curr") / 1000.0
    profit_after_tax = (g2(2350, "curr") - g2(2355, "curr")) / 1000.0
    tax_charge = -g2(2300, "curr") / 1000.0
    profit_before_tax = profit_after_tax + tax_charge

    extra_writeoffs = -g2(2240, "curr") / 1000.0
    extra_other = g2(2270, "curr") / 1000.0
    extra_items = extra_writeoffs + extra_other

    bank_int_income = -g2(2220, "curr") / 1000.0
    bank_int_expense = g2(2250, "curr") / 1000.0
    net_bank_interest = bank_int_income + bank_int_expense

    ebit = profit_before_tax + extra_items + net_bank_interest
    ebit_pct = ebit / turnover if turnover else 0.0

    depreciation = g2(2515, "curr") / 1000.0
    ebitda = ebit + depreciation
    ebitda_pct = ebitda / turnover if turnover else 0.0

    ebda = ebitda - tax_charge - net_bank_interest
    ebda_pct = ebda / turnover if turnover else 0.0

    # 3. New investments
    new_inv_tangible = (
        g1(1011, "curr") - g1(1011, "prev")
    ) / 1000.0 + max(0.0, (g1(1005, "curr") - g1(1005, "prev")) / 1000.0)

    # 4. Gearing
    equity = g1(1495, "curr") / 1000.0
    bank_loans = (
        (g1(1510, "curr") + g1(1600, "curr"))
        if (g1(1510, "curr") + g1(1600, "curr")) != 0
        else g1(1595, "curr")
    ) / 1000.0
    other_loans = (g1(1515, "curr") + g1(1610, "curr")) / 1000.0
    interest_debt = bank_loans + other_loans
    cash_in_bank = g1(1165, "curr") / 1000.0
    total_assets = g1(1300, "curr") / 1000.0
    equity_to_assets = equity / total_assets if total_assets else 0.0

    # 5. Efficiency Ratios
    trade_creditors = g1(1615, "curr") / 1000.0
    trade_debtors = g1(1125, "curr") / 1000.0
    inventory = g1(1100, "curr") / 1000.0

    # 6. Liquidity Ratios
    current_assets = g1(1195, "curr") / 1000.0
    current_liab = g1(1695, "curr") / 1000.0

    # 7. Other Ratios
    payroll_cost = (g2(2505, "curr") + g2(2510, "curr")) / 1000.0

    return {
        # Словник для відображення
        "top_box": [
            ("Exchange rate UAH/EUR (input)", f"{eur_rate:.2f}"),
            ("Fixed assets, MIO EUR", f"{fixed_assets_eur:.2f}"),
            ("Fixed assets % in total ASSETS", f"{fixed_assets_pct * 100:.1f}%"),
            ("Turnover growth YoY", f"{turnover_growth * 100:.1f}%"),
            (
                "Investment real estate, MIO EUR",
                f"{inv_real_estate:.2f}" if inv_real_estate > 0 else "-",
            ),
        ],
        "bridge_table": [
            ("TURNOVER", f"{turnover:.2f}", "row_bold"),
            ("Profit (+)/Loss (-) AFTER tax", f"{profit_after_tax:.2f}", ""),
            (
                "Add Back Tax Charge (+)/Credit (-)",
                f"{tax_charge:.2f}" if tax_charge else "-",
                "",
            ),
            (
                "Profit (+)/Loss (-) before tax",
                f"{profit_before_tax:.2f}",
                "row_bold",
            ),
            (
                "Extraordinary items",
                f"{extra_items:.2f}" if extra_items else "-",
                "row_gray",
            ),
            (
                "Currency translation difference",
                "-",
                "row_gray",
            ),
            (
                "Net interest & similar charges - Bank interest",
                f"{net_bank_interest:.2f}" if net_bank_interest else "-",
                "row_gray",
            ),
            ("EBIT", f"{ebit:.2f}", "row_gray_bold"),
            ("As percentage of turnover", f"{ebit_pct * 100:.1f}%", ""),
            ("Depreciation and Amortisation", f"{depreciation:.2f}", ""),
            (
                "Depreciation, Amortisation & Grants",
                f"{depreciation:.2f}",
                "row_gray_bold",
            ),
            ("EBITDA", f"{ebitda:.2f}", "row_gray_bold"),
            ("As percentage of turnover", f"{ebitda_pct * 100:.1f}%", ""),
            (
                "Less Tax Charge (+)/Credit (-)",
                f"{tax_charge:.2f}" if tax_charge else "-",
                "",
            ),
            ("EBDA", f"{ebda:.2f}", "row_gray_bold"),
            ("As percentage of turnover", f"{ebda_pct * 100:.1f}%", ""),
            ("New investments - tangible", f"{new_inv_tangible:.2f}", ""),
            ("New investments", f"{new_inv_tangible:.2f}", "row_gray_bold"),
            ("Equity", f"{equity:.2f}", ""),
            (
                "Equity (total shareholders funds)",
                f"{equity:.2f}",
                "row_gray_bold",
            ),
            (
                "Interest bearing debt",
                f"{interest_debt:.2f}" if interest_debt else "-",
                "row_gray_bold",
            ),
            ("cash in bank (+)", f"{cash_in_bank:.2f}", ""),
            ("Total Assets", f"{total_assets:.2f}", "row_gray_bold"),
            (
                "Equity / Total assets",
                f"{equity_to_assets * 100:.1f}%",
                "row_bold",
            ),
            ("Trade Creditors", f"{trade_creditors:.2f}", ""),
            ("Trade Debtors", f"{trade_debtors:.2f}", ""),
            ("Inventory", f"{inventory:.2f}", ""),
            (
                "Current Assets (from financial statements)",
                f"{current_assets:.2f}",
                "",
            ),
            ("Current Assets", f"{current_assets:.2f}", "row_gray_bold"),
            (
                "Current Liabilities (from financial statements)",
                f"{current_liab:.2f}",
                "",
            ),
            ("Current Liabilities", f"{current_liab:.2f}", "row_gray_bold"),
            ("Payroll cost", f"{payroll_cost:.2f}", ""),
        ],
        # Сирі значення для запису в Excel
        "raw": {
            "eur_rate": eur_rate,
            "fixed_assets_eur": fixed_assets_eur,
            "fixed_assets_pct": fixed_assets_pct,
            "turnover_growth": turnover_growth,
            "turnover": turnover,
            "profit_after_tax": profit_after_tax,
            "tax_charge": tax_charge,
            "profit_before_tax": profit_before_tax,
            "ebit": ebit,
            "ebit_pct": ebit_pct,
            "depreciation": depreciation,
            "ebitda": ebitda,
            "ebitda_pct": ebitda_pct,
            "ebda": ebda,
            "ebda_pct": ebda_pct,
            "new_inv": new_inv_tangible,
            "equity": equity,
            "cash": cash_in_bank,
            "total_assets": total_assets,
            "equity_to_assets": equity_to_assets,
            "trade_creditors": trade_creditors,
            "trade_debtors": trade_debtors,
            "inventory": inventory,
            "current_assets": current_assets,
            "current_liab": current_liab,
            "payroll_cost": payroll_cost,
        },
    }


def update_excel_kfi_sheet(
    original_bytes: bytes, kfi_results: dict
) -> bytes:
    """Оновлює оригінальний файл (усі 4 вкладки залишаються), проставляючи значення на 4-му аркуші KFI."""
    wb = openpyxl.load_workbook(io.BytesIO(original_bytes))

    if "KFI" not in wb.sheetnames:
        raise ValueError("У книзі відсутній 4-й аркуш з назвою 'KFI'.")

    ws = wb["KFI"]
    raw = kfi_results["raw"]

    # Оновлюємо курс у комірці H1
    ws["H1"] = raw["eur_rate"]

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()