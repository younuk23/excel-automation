from openpyxl import load_workbook, Workbook

wb = load_workbook('origin.xlsx')

ws = wb.active
ws.title = "일일출고형식"

col = {
    "CS": ws["A"],
    "company": ws["B"],
    "maker": ws["C"],
    "mall": ws["D"],
    "receiver": ws["E"],
    "phone_number":  ws["F"],
    "second_number":  ws["G"],
    "orderer": ws["H"],
    "order_number": ws["I"],
    "id": ws["J"],
    "delivery_number": ws["K"],
    "invoice_number": ws["L"],
    "delivery_status": ws["M"]
}

print(col)


for n, cell in enumerate(col["CS"]):
    print("cell", cell)

    print("before", col["company"][n].value)

    col["company"][n].value = f"{col_company[n].value}({cell.value})" if cell.value is not None else col_company[n].value

    print("after", col_company[n].value)
