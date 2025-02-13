import argparse

import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime


def _convert_date(date):
    return pd.to_datetime(date).date().strftime('%Y-%m-%d')


def _convert_decimal(value):
    return str(value).replace(",", ".")

def generate_header(input_file_path):
    header_sheet_name = "Podatki"
    header_df = pd.read_excel(input_file_path, sheet_name=header_sheet_name)
    header_data = header_df.iloc[0]

    envelope = ET.Element("Envelope",
                          xmlns="http://edavki.durs.si/Documents/Schemas/Doh_KDVP_9.xsd",
                          attrib={
                              "xmlns:edp": "http://edavki.durs.si/Documents/Schemas/EDP-Common-1.xsd"
                          }
                          )

    header = ET.SubElement(envelope, "edp:Header")
    taxpayer = ET.SubElement(header, "edp:taxpayer")

    for field in header_df.columns:
        ET.SubElement(taxpayer, f"edp:{field}").text = str(header_data[field])

    ET.SubElement(envelope, "edp:AttachmentList")
    ET.SubElement(envelope, "edp:Signatures")

    return envelope

def generate_kdvp_element(input_file_path):
    body = ET.Element("body")
    ET.SubElement(body, "edp:bodyContent")

    doh_kdvp = ET.SubElement(body, "Doh_KDVP")
    kdvp = ET.SubElement(doh_kdvp, "KDVP")

    kdvp_df = pd.read_excel(input_file_path, sheet_name="KDVP podatki", dtype={"TelephoneNumber": str, "IsResident": bool})
    kdvp_data = kdvp_df.iloc[0]

    year = kdvp_data["Year"]
    is_resident = kdvp_data["IsResident"]
    telephone_number = kdvp_data["TelephoneNumber"]
    email = kdvp_data["Email"]

    ET.SubElement(kdvp, "DocumentWorkflowID").text = "0"
    ET.SubElement(kdvp, "Year").text = str(year)
    ET.SubElement(kdvp, "PeriodStart").text = f"{year}-01-01"
    ET.SubElement(kdvp, "PeriodEnd").text = f"{year}-12-31"
    ET.SubElement(kdvp, "IsResident").text = str(is_resident).lower()
    ET.SubElement(kdvp, "TelephoneNumber").text = str(telephone_number)

    sheet_names = pd.ExcelFile(input_file_path).sheet_names
    securities_count = len([sheet for sheet in sheet_names if sheet not in ["KDVP podatki", "Podatki"]])
    ET.SubElement(kdvp, "SecurityCount").text = str(securities_count)

    hardcoded_fields = {
        "SecurityShortCount": "0",
        "SecurityWithContractCount": "0",
        "SecurityWithContractShortCount": "0",
        "ShareCount": "0",
        "SecurityCapitalReductionCount": "0",
    }

    for field, value in hardcoded_fields.items():
        ET.SubElement(kdvp, field).text = str(value)

    ET.SubElement(kdvp, "Email").text = str(email)

    return body, doh_kdvp

def generate_kdvp_item_element(sheet_name, doh_kdvp_element):
    kdvp_item_df = pd.read_excel(arguments.xlsx_input, sheet_name=sheet_name)

    # Create the KDVPItem element
    kdvp_item = ET.SubElement(doh_kdvp_element, "KDVPItem")

    # Add the fields that should be directly under KDVPItem
    ET.SubElement(kdvp_item, "InventoryListType").text = "PLVP"
    ET.SubElement(kdvp_item, "Name").text = sheet_name

    boolean_columns = ["HasForeignTax", "HasLossTransfer", "ForeignTransfer", "TaxDecreaseConformance", "IsFond"]
    for col in boolean_columns:
        kdvp_item_df[col] = kdvp_item_df[col].astype(bool)

    # Create the XML elements for boolean fields under KDVPItem
    ET.SubElement(kdvp_item, "HasForeignTax").text = str(kdvp_item_df.iloc[0]["HasForeignTax"]).lower()
    ET.SubElement(kdvp_item, "HasLossTransfer").text = str(kdvp_item_df.iloc[0]["HasLossTransfer"]).lower()
    ET.SubElement(kdvp_item, "ForeignTransfer").text = str(kdvp_item_df.iloc[0]["ForeignTransfer"]).lower()
    ET.SubElement(kdvp_item, "TaxDecreaseConformance").text = str(kdvp_item_df.iloc[0]["TaxDecreaseConformance"]).lower()

    # Create the Securities element under KDVPItem
    securities = ET.SubElement(kdvp_item, "Securities")
    ET.SubElement(securities, "ISIN").text = sheet_name
    ET.SubElement(securities, "IsFond").text = str(kdvp_item_df.iloc[0]["IsFond"]).lower()

    # Iterate over rows and add individual security rows under Securities
    for idx, row in kdvp_item_df.iterrows():
        row_element = ET.SubElement(securities, "Row")
        ET.SubElement(row_element, "ID").text = str(idx)

        if row["Type"] == "B":
            purchase = ET.SubElement(row_element, "Purchase")
            ET.SubElement(purchase, "F1").text = _convert_date(row["Date"])
            ET.SubElement(purchase, "F2").text = "B"
            ET.SubElement(purchase, "F3").text = _convert_decimal(str(row["Quantity"]))
            ET.SubElement(purchase, "F4").text = _convert_decimal(str(row["Price"]))
            ET.SubElement(purchase, "F5").text = _convert_decimal(str(row["GiftAndInheritanceTax"]))
        elif row["Type"] == "S":
            sale = ET.SubElement(row_element, "Sale")
            ET.SubElement(sale, "F6").text = _convert_date(row["Date"])
            ET.SubElement(sale, "F7").text = _convert_decimal(str(row["Quantity"]))
            ET.SubElement(sale, "F9").text = _convert_decimal(str(row["Price"]))

        ET.SubElement(row_element, "F8").text = _convert_decimal(str(row["Remaining"]))

    return kdvp_item


def parse_arguments():
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--xlsx_input",
        type=str,
        required=True,
        help="Input XLSX file path",
    )
    parser.add_argument(
        "--xml_output",
        type=str,
        required=False,
        help="Output XML file path",
    )
    input_args = parser.parse_args()
    return input_args


if __name__ == "__main__":
    arguments = parse_arguments()
    header = generate_header(arguments.xlsx_input)
    kdvp_element = generate_kdvp_element(arguments.xlsx_input)
    header.append(kdvp_element[0])

    xlsx_file = pd.ExcelFile(arguments.xlsx_input)

    processed_sheets = set()

    for sheet in xlsx_file.sheet_names:
        print("Iteration")
        if sheet not in ["KDVP podatki", "Podatki"] and sheet not in processed_sheets:
            kdvp_item = generate_kdvp_item_element(sheet, kdvp_element[1])
            processed_sheets.add(sheet)

    tree = ET.ElementTree(header)
    ET.indent(tree, space="    ", level=0)
    ET.dump(header)

    if arguments.xml_output:
        tree.write(arguments.xml_output, encoding="utf-8", xml_declaration=True)
