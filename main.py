import collections
import collections.abc
import json
import os
import pandas as pd
from powerpoint import Powerpoint

# IMPORTANT: MUST NAME SLIDE SHAPES https://www.youtube.com/watch?v=IhES3of_9Nw

# Load config file
with open("config.json", "r") as config_file:
    config = json.load(config_file)
# Make sure it's valid
assert(os.path.exists(config.get("template_path")))
assert(os.path.exists(config.get("excel_path")))

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
        pptx.set_chart_data(slide, chart_name, values)

def update_other(pptx: Powerpoint, slide, provider_data : dict) -> None:
    for element_name, function_name in OTHER.items():
        pptx.update_other(slide, element_name, provider_data, function_name)

def update_text(pptx : Powerpoint, slide, provider_data) -> None:
    for element_name, fstring in ELEMENT_TO_FSTRING.items():
        text = fstring.format(**provider_data)
        pptx.update_text(slide, element_name, text)

def generate_slide(pptx, slide, provider):
    print(f"Generating slide for provider '{provider}'")
    provider_data = search_excel_sheet(config.get("excel_path"), config.get("sheet_name"),
                                       config.get("header_row") - 1, config.get("target_column"), provider)

    # Manipulate slide
    print("Updating text objects")
    update_text(pptx, slide, provider_data)
    print("Updating charts")   
    update_charts(pptx, slide, provider_data)
    print("Updating other")
    update_other(pptx, slide, provider_data)
    print(f"Completed slide for provider '{provider}'")

def generate_all_slides():
    # get all slides
    all_excel_data = pd.read_excel(config.get("excel_path"), sheet_name=config.get("sheet_name"),
                                   header=config.get("header_row") - 1)
    all_targets = list(all_excel_data[config.get("target_column")].unique())
    # New slides get inserted at index 1; iterate in reverse to get proper order
    all_targets.sort(reverse=True) 
    del all_excel_data
    print(f"Found {len(all_targets)} slides to generate")
    # Create powerpoint
    pptx = Powerpoint(config.get("template_path"), config.get("output_path"))
    for target in all_targets:
        slide = pptx.new_slide()
        generate_slide(pptx, slide, target)
    pptx.close()

generate_all_slides()
