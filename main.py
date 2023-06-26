import pretty_errors # TODO REMOVE

import collections
import collections.abc
import json
import pandas as pd
from powerpoint import Powerpoint

# IMPORTANT: MUST NAME SLIDE SHAPES https://www.youtube.com/watch?v=IhES3of_9Nw

EXCEL_FILEPATH = "//ACM4/atlanticfiles/Projects/2023 files/EdgeConneX Colo Mystery Shopping - 23024/Cal_verify_quotes_sterilized.xlsx"
PPTX_PATH = r"C:/Users/cnightingale/excel2slides/template_slide.pptx"
EXCEL_SHEET_NAME = "Raw Data"
HEADER_ROW = 3
TARGET_COLUMN = "Provider"
SEARCH_TERM = "Cyrus"

with open("columns.json", "r") as columns_file:
    RELEVANT_COLUMNS = json.load(columns_file)
with open("text.json", "r") as text_file:
    ELEMENT_TO_FSTRING = json.load(text_file)
with open("charts.json", "r") as charts_file:
    CHARTS = json.load(charts_file)
with open("other.json", "r") as other_file:
    OTHER = json.load(other_file)

def search_excel_sheet(filepath : str, sheet : str, header_row : int, target_column : str, search_term : str) -> pd.DataFrame:
    # Read the Excel file into a DataFrame
    df = pd.read_excel(filepath, sheet_name=sheet, header=header_row)
    # Get all rows that correspond to the search term
    returned_rows = df[df[target_column].str.contains(search_term)]
    # Keep only the specified columns
    relevant_data = returned_rows[list(RELEVANT_COLUMNS.keys())]
    # Rename the columns as specified
    named_data = relevant_data.rename(columns=RELEVANT_COLUMNS)
    # Convert DataFrame to dictionary
    dict_data = named_data.to_dict(orient='list')
    # Check for columns with a single unique value
    for column_name, values in dict_data.items():
        unique_values = set(values)
        if len(unique_values) == 1:
            dict_data[column_name] = [unique_values.pop()]

    return dict_data

def find_shape_in_group(group, shape_name):
    for shape in group.shapes:
        if shape.shape_type == 6:  # 6 represents a grouped shape
            found_shape = find_shape_in_group(shape, shape_name)
            if found_shape:
                return found_shape
        elif shape.name == shape_name:
            return shape

def get_shape_by_name(slide, shape_name):
    # TODO elegantly raise exception when shape is missing? Can't just put in recursive method cuz it will never find
    # first check placholders     
    for shape in slide.placeholders:
        if shape.name == shape_name:
            return shape
    # If not found, now check shapes (recursively to check groups)
    return find_shape_in_group(slide, shape_name)

def update_charts(pptx : Powerpoint, slide, provider_data : dict) -> None:
    for chart_name, data_cols in CHARTS.items():
        # Retrieve data for chart as per charts.json
        values = [provider_data.get(data_col) for data_col in data_cols]
        # Un-collapse data if values for different categories were identical
        max_len = max([len(col) for col in values])
        for i, value in enumerate(values):
            if len(value) == 1:
                values[i] = value * max_len
        pptx.set_chart_data(PPTX_PATH, slide, chart_name, values)

def update_other(pptx: Powerpoint, slide, provider_data : dict) -> None:
    for element_name, function_name in OTHER.items():
        pptx.update_other(slide, element_name, provider_data, function_name)

# TODO Add slide duplication when they resolve this git issue https://github.com/scanny/python-pptx/issues/132
# Until then, workaround is to manually copy/paste the template slide n times
def update_text(pptx : Powerpoint, slide, provider_data) -> None:
    for element_name, fstring in ELEMENT_TO_FSTRING.items():
        text = fstring.format(**provider_data)
        pptx.update_text(slide, element_name, text)

def generate_slide(pptx, slide, provider):
    print(f"Generating slide for provider '{provider}'")
    provider_data = search_excel_sheet(EXCEL_FILEPATH, EXCEL_SHEET_NAME, HEADER_ROW, TARGET_COLUMN, provider)

    # Manipulate slide
    print("Updating text objects")
    update_text(pptx, slide, provider_data)
    print("Updating charts")   
    update_charts(pptx, slide, provider_data)
    print("Updating other")
    update_other(pptx, slide, provider_data)
    print(f"Completed slide for provider '{provider}'")

def generate_all_slides():
    # get all providers
    all_excel_data = pd.read_excel(EXCEL_FILEPATH, sheet_name=EXCEL_SHEET_NAME, header=HEADER_ROW)
    all_providers = all_excel_data['Provider'].unique()
    all_providers.sort()
    del all_excel_data
    print(f"Found {len(all_providers)} slides to generate")
    # Create powerpoint
    pptx = Powerpoint(r"C:/Users/cnightingale/excel2slides/template_slide.pptx")
    for provider in all_providers:
        slide = pptx.new_slide()
        generate_slide(pptx, slide, provider)
    pptx.close()


generate_all_slides()
