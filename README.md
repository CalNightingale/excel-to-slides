# excel-to-slides
# Setup
## Step 1: Install git
Open your terminal, type `git --version`, and hit enter. If you have git installed already, this will print a version number.
If not, you will need to [install git](https://git-scm.com/downloads)
## Step 2: Install python
Open your terminal, type `python --version`, and hit enter. If you have python installed already, this will print a version number.
If not, you will need to [install python](https://www.python.org/downloads/)
## Step 3: Create a virtual environment
This should be done locally, not on the server.
By default your terminal will likely open to `C:\Users\your_username`, which is a fine location for your virtual environment.
Type `python -m venv my_env` and hit enter
## Step 4: Activate your virtual environment
On windows, this can be done by running `.\my_env\Scripts\activate`.
You can verify this works by checking that your terminal reads `(my_env) C:\Users\your-username>` in the bottom left
## Step 5: Clone the program code
Run the command `git clone https://github.com/CalNightingale/excel-to-slides.git`.
This will download the code into a new folder called `excel-to-slides`
## Step 6: Enter the code folder
In the windows terminal, this is done with `chdir excel-to-slides`.
You can verify this works by checking that your terminal reads `(my_env) C:\Users\your-username\excel-to-slides>`
## Step 7: Install required libraries
Run the command `pip install -r requirements.txt`.
This will install libraries that the code needs in order to function properly. This may take a minute or two

# How to Use
## Step 1: Create template slide
Ideally you can begin by copying in an existing slide that is formatted properly.
Create such a slide if you are starting from scratch.
It is very important that you NAME THE ELEMENTS (text boxes, charts, etc.) OF YOUR SLIDE.
Follow [this video](https://www.youtube.com/watch?v=IhES3of_9Nw) for a guide.
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

## Step 4: Create `text.json` file
Create a file called `text.json` that maps the powerpoint element names you created in Step 1 to the text you want to write in those elements.
Skip elements like charts that you do not want to put text in.
Place values that should be pulled from the excel file in {brackets}, and use the column names you created in step 3.
```json
{
    "title" : "Provider Detail - {provider}",
    "header": "{provider} quoted {prod_type} for {mkt}"
}
```
## Step 5: Create `charts.json` file
Create a file called `charts.json` that maps the chart names you specified in Step 1 to the columns of data that they should reference.
Use the column names you created in step 3. An example is below
```json
{
    "mrc_chart" : ["mkt", "mrc_kw"],
    "nrc_chart" : ["mkt", "nrc_cab"]
}
```

## Step 6: Create `other.json` file
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
Once everything is ready, just run `python .\main.py`. Make sure you're in the virtual environment you made in (see Setup Step 4).
