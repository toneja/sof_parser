#!/usr/bin/python3

import os
import re

import openpyxl
import pdfplumber
import pandas as pd
from tkinter import Tk, filedialog


def clean_money(value):
    return float(value.replace(",", ""))


def build_lines(page):
    words = page.extract_words(x_tolerance=2, y_tolerance=2, keep_blank_chars=False)
    rows = {}
    for word in words:
        y = round(word["top"])
        rows.setdefault(y, []).append(word)
    lines = []
    for y in sorted(rows):
        row = sorted(rows[y], key=lambda w: w["x0"])
        text = " ".join(word["text"] for word in row)
        lines.append(text)
    return lines


def parse_rows(lines):
    # parse the rows
    rows = []
    for line in lines:
        tokens = line.split()
        if len(tokens) < 7:
            continue
        first_token = tokens[0]
        # we only want rows with object codes
        if not re.compile(r"^\d{4}$").match(first_token):
            continue
        numeric_tokens = [token for token in tokens if re.compile(r"^-?\d[\d,]*\.\d{2}$").match(token)]
        if len(numeric_tokens) < 5:
            continue
        try:
            financial_plan = clean_money(numeric_tokens[-5])
            reconciled = clean_money(numeric_tokens[-4])
            unreconciled = clean_money(numeric_tokens[-3])
            total_obligations = clean_money(numeric_tokens[-2])
            balance_available = clean_money(numeric_tokens[-1])
        except ValueError:
            continue
        description_tokens = []
        for token in tokens[1:]:
            if token == numeric_tokens[-5]:
                break
            description_tokens.append(token)
        description = " ".join(description_tokens)
        rows.append(
            {
                "ObjectCode": int(first_token),
                "Description": description,
                "FinancialPlan": financial_plan,
                "Reconciled": reconciled,
                "Unreconciled": unreconciled,
                "TotalObligations": total_obligations,
                "BalanceAvailable": balance_available,
            }
        )
    return rows


def parse_pdf(pdf_file):
    # extract text from "STATUS OF FUNDS" pages
    index = 0
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            lines = build_lines(page)
            page_text = "\n".join(lines).upper()
            if (
                "STATUS OF FUNDS" in page_text
                and "COMMENTS" not in page_text
                and "EXPIRED" not in page_text
            ):
                # write data to the output file
                df = pd.DataFrame(parse_rows(lines))
                if (
                    index == 0
                ):  # First sheet is always the Totals, write it to a new file
                    with pd.ExcelWriter(
                        pdf_file.replace(".pdf", "-PARSED.xlsx"), engine="openpyxl"
                    ) as writer:
                        df.to_excel(writer, sheet_name=f"Account Summary", index=False)
                        index += 1
                else:
                    with pd.ExcelWriter(
                        pdf_file.replace(".pdf", "-PARSED.xlsx"),
                        engine="openpyxl",
                        mode="a",
                    ) as writer:
                        df.to_excel(
                            writer, sheet_name=f"SubAccount {index}", index=False
                        )
                        index += 1


def parse_xlsx(xlsx_file):
    # Sub Accounts
    sub_accounts = {
        "001": "Jana,Choi",
        "045": "Mahaffee,Stockwell,Grunwald",
        "046": "Weiland,Zasada,Mollov",
        "129": "IR4 Weiland",
    }
    df = pd.read_excel(xlsx_file)
    for account in sub_accounts.keys():
        sub_df = df[df["Detail Sub Account"] == int(account)].copy()
        # sort data by Object Code
        sub_df = sub_df.sort_values(by="Object Class", key=lambda x: x.astype(str).str[:3].astype(int))
        with pd.ExcelWriter(
            xlsx_file.replace(".xlsx", "-PARSED.xlsx"),
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace",
        ) as writer:
            sub_df.to_excel(writer, sheet_name=sub_accounts.get(account), index=False)


def format_workbook(
    filename,
    currency_columns=None,
    column_padding=3,
    currency_format="$#,##0.00",
):
    if currency_columns is None:
        currency_columns = []
    wb = openpyxl.load_workbook(filename)
    green_fill = openpyxl.styles.PatternFill(
        fill_type="solid", start_color="C6EFCE", end_color="C6EFCE"
    )
    yellow_fill = openpyxl.styles.PatternFill(
        fill_type="solid", start_color="FFEB9C", end_color="FFEB9C"
    )
    red_fill = openpyxl.styles.PatternFill(
        fill_type="solid", start_color="FFC7CE", end_color="FFC7CE"
    )
    for ws in wb.worksheets:
        max_row = ws.max_row
        max_col = ws.max_column
        for col in range(1, max_col + 1):
            header = ws.cell(row=1, column=col).value
            if header is None:
                header = ""
            header = str(header)
            column_letter = openpyxl.utils.get_column_letter(col)
            # Fit width to header
            ws.column_dimensions[column_letter].width = len(header) + column_padding
            if header not in currency_columns:
                continue
            cell_range = f"{column_letter}2:{column_letter}{max_row}"
            # Apply currency number format
            for row in range(2, max_row + 1):
                ws.cell(row=row, column=col).number_format = currency_format
            # add a totals row
            total_row = max_row + 1
            ws[f"A{total_row}"] = "Total"
            ws[f"A{total_row}"].font = openpyxl.styles.Font(bold=True)
            # SUM formula
            total_cell = ws.cell(row=total_row, column=col)
            total_cell.value = f"=SUM({column_letter}2:{column_letter}{max_row})"
            total_cell.number_format = currency_format
            total_cell.font = openpyxl.styles.Font(bold=True)
            if not "Account" in ws.title:
                continue
            # simple conditional formatting
            ws.conditional_formatting.add(
                cell_range,
                openpyxl.formatting.rule.CellIsRule(
                    operator="lessThan",
                    formula=["0"],
                    fill=red_fill,
                ),
            )
            ws.conditional_formatting.add(
                cell_range,
                openpyxl.formatting.rule.CellIsRule(
                    operator="equal",
                    formula=["0"],
                    fill=yellow_fill,
                ),
            )
            ws.conditional_formatting.add(
                cell_range,
                openpyxl.formatting.rule.CellIsRule(
                    operator="greaterThan",
                    formula=["0"],
                    fill=green_fill,
                ),
            )
    wb.save(filename)


if __name__ == "__main__":
    os.chdir(os.path.dirname(__file__))
    root = Tk()
    root.withdraw()
    pdf_file = filedialog.askopenfilename(
        initialdir=".",
        title="Select Status of Funds PDF",
        filetypes=[("PDF Files", "*.pdf")],
    )
    if not os.path.exists(pdf_file):
        quit(f"{pdf_file} not found.")
    xlsx_file = pdf_file.replace(".pdf", ".xlsx")
    if not os.path.exists(xlsx_file):
        quit(f"{xlsx_file} not found.")
    # parse the files
    parse_pdf(pdf_file)
    parse_xlsx(xlsx_file)
    # format the new workbook
    format_workbook(
        xlsx_file.replace(".xlsx", "-PARSED.xlsx"),
        currency_columns=[
            "FinancialPlan",
            "Reconciled",
            "Unreconciled",
            "TotalObligations",
            "BalanceAvailable",
            "Unreconciled Amt",
            "Reconciled Amt",
        ],
    )
