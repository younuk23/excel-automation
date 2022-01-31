from openpyxl import load_workbook
from enum import Enum

wb = load_workbook('test.xlsx')

ws = wb.active
ws.title = "일일출고형식"

columns = {
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
    "delivery_status": ws["M"],
    "product_name": ws["N"],
    "option": ws["O"],
    "quantity": ws["P"],
    "box": ws["Q"],
    "regional": ws["R"],
    "address": ws["S"],
    "delivery_company": ws["T"],
    "prepaid": ws["U"],
    "pay_on_delivery": ws["V"],
    "product_cost": ws["W"],
    "delivery_man_cost": ws["X"],
    "order_enrollment_day": ws["Y"],
    "warehousing_day": ws["Z"],
    "designated_day": ws["AA"],
    "departure_from_branch_store_day": ws["AB"],
    "arrived_at_branch_store": ws["AC"],
    "delivery_start": ws["AD"],
    "delivery_complete": ws["AE"],
    "recovery": ws["AF"],
    "delivery_message": ws["AG"],
    "memo": ws["AH"],
    "head_office_memo": ws["AI"],
    "modified_man": ws["AJ"]
}

def merge_CS_to_company():
    for n, cell in enumerate(columns["CS"]):
        if n == 0:
            continue

        origin_value = columns["company"][n].value

        columns["company"][n].value = (
                f"{origin_value}({cell.value})"
                if cell.value is not None
                else origin_value
                )


def change_column_title_for_box():
    columns["box"][0].value = "수량"




def write_necessary_columns():
    ws.insert_cols(0, 11)

    necessary_columns = [
        columns["company"],
        columns["product_name"],
        columns["option"],
        columns["box"],
        columns["receiver"],
        columns["phone_number"],
        columns["second_number"],
        columns["address"],
        columns["pay_on_delivery"],
        columns["prepaid"],
        columns["delivery_message"]
        ]

    for col in ws.iter_cols(min_col=1, max_col=11):
        for cell in col:
            source = necessary_columns[cell.column - 1]
            cell.value = source[cell.row - 1].value

    ws.delete_cols(12, 99)

def init():
    merge_CS_to_company()
    change_column_title_for_box()
    write_necessary_columns()


init()


wb.save("test_result.xlsx")

