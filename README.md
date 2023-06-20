# excel-to-slides
# How to Use
## Step 1: Create powerpoint file
Ideally you can begin by copying in an existing slide that is formatted properly.
Create such a slide if you are starting from scratch.
It is very important that you NAME THE ELEMENTS (text boxes, charts, etc.) OF YOUR SLIDE.
Follow this video for a guide: https://www.youtube.com/watch?v=IhES3of_9Nw

Once you have created your template slide, you must MANUALLY COPY/PASTE IT FOR EACH SLIDE YOU PLAN ON FORMATTING.
This seems rediculous but the library this code relies on actuallly cannot duplicate slides for some reason.
See this git issue for more info: https://github.com/scanny/python-pptx/issues/132

## Step 2: Create columns.json file
Create a file called `columns.json` that maps the names of the columns you want to keep in the excel to more code-friendly names (no spaces, etc.).
An example is below:
```json
{
    "Provider Name" : "provider",
    "Product Description" : "prod_type",
    "Market" : "mkt"
}
```

## Step 3: Create text.json file
Create a file called `text.json` that maps the powerpoint element names you created in Step 1 to the text you want to write in those elements.
Skip elements like charts that you do not want to put text in.
Place values that should be pulled from the excel file in {brackets}, and use the column names you created in step 2.
```json
{
    "title" : "Provider Detail - {provider}",
    "header": "{provider} quoted {prod_type} for {mkt}"
}
```
