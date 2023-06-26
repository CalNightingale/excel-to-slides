# excel-to-slides
# How to Use
## Step 1: Create powerpoint file
Ideally you can begin by copying in an existing slide that is formatted properly.
Create such a slide if you are starting from scratch.
It is very important that you NAME THE ELEMENTS (text boxes, charts, etc.) OF YOUR SLIDE.
Follow this video for a guide: https://www.youtube.com/watch?v=IhES3of_9Nw
You do not have to name all elements, only the ones that you would like the script to modify

## Step 2: Create a config.json file
Create a file called `config.json` that defines the following properties:
```json
{
    "output_path" : "path/to/output.pptx",
    "template_path" : "path/to/template.pptx",
    "excel_path" : "path/to/excel/data.xlsx",
    "sheet_name" : "Name of Relevant Sheet",
    "header_row" : 0,
    "target_column" : "Name of Target Column in Excel"
}
```
Many of these are self explanatory, but header row means the row number in excel that corresonds to the titles of the data.
Target column means the column of data for which each unique value should have it's own slide. In the example of a mystery shopping case,
this would be `"Provider"`

## Step 3: Create columns.json file
Create a file called `columns.json` that maps the names of the columns you want to keep in the excel to more code-friendly names (no spaces, etc.).
An example is below:
```json
{
    "Provider Name" : "provider",
    "Product Description" : "prod_type",
    "Market" : "mkt"
}
```

## Step 4: Create text.json file
Create a file called `text.json` that maps the powerpoint element names you created in Step 1 to the text you want to write in those elements.
Skip elements like charts that you do not want to put text in.
Place values that should be pulled from the excel file in {brackets}, and use the column names you created in step 3.
```json
{
    "title" : "Provider Detail - {provider}",
    "header": "{provider} quoted {prod_type} for {mkt}"
}
```
## Step 5: Create charts.json file
Create a file called `charts.json` that maps the chart names you specified in Step 1 to the columns of data that they should reference.
Use the column names you created in step 3. An example is below
```json
{
    "mrc_chart" : ["mkt", "mrc_kw"],
    "nrc_chart" : ["mkt", "nrc_cab"]
}
```

## Step 6: Create other.json file
Create a file called `other.json` that maps the names of non-text, non-chart elements from Step 1 to the names of functions in `utils.py`
that can be used to update them. This should be used for slide elements that need some sort of custom update rule and are not as
simple as just changing the text or the chart data. An example would be a map of which markets a provider is available in (see below);
this would need to map to a function that generates the map for the given provider on each slide.
```json
{
    "market_table" : "handle_mkt_presence_table",
    "market_map" : "handle_mkt_map"
}
```
It is likely that you may need to write a new function in `utils.py` when trying to automate a new type of deck. These functions should be
speficied as follows:
```python
def handle_ELEMENT_NAME_HERE(slide, element, provider_data):
    # Code goes here
```
Note that functions MUST take these three arguments, in this order. They will cause errors otherwise! See the existing functions already in
`utils.py` for reference

## Step 7: Run the script
### Install python
Open powershell and type `python --version` to see whether you have python installed. If you don't you will need to
[install it](https://www.python.org/downloads/)
### Create a virtual environment
This should be done locally, not on the server. You can create one in your home directory with
`python -m venv ~/venv`
### Activate virtual environment
`~\venv\Scripts\activate`
### Install required libraries
`pip install -r requirements.txt`
### Run it!
`python .\main.py`
