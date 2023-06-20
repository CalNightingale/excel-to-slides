# excel-to-slides
# How to Use
## Step 1: Create powerpoint file
Ideally you can begin by copying in an existing slide that is formatted properly.
Create such a slide if you are starting from scratch.
It is very important that you NAME THE ELEMENTS OF YOUR SLIDE.
Follow this video for a guide: https://www.youtube.com/watch?v=IhES3of_9Nw

Once you have created your template slide, you must MANUALLY COPY/PASTE IT FOR EACH SLIDE YOU PLAN ON FORMATTING.
This seems rediculous but the library this code relies on actuallly cannot duplicate slides for some reason.
See this git issue for more info: https://github.com/scanny/python-pptx/issues/132

## Step 2: Create columns.json file
Create a file `columns.json` that maps the names of the columns you want to keep in the excel to more code-friendly names (no spaces, etc.).
An example is below:
```json
{
    "Excel Column 1" : "column_1",
    "Excel Column 2" : "column_2
}
```
