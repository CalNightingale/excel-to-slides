import win32com.client

SHAPE_NAME = "mrc_chart"

# Create an instance of the PowerPoint application
powerpoint = win32com.client.Dispatch("PowerPoint.Application")

# Set to invisible
powerpoint.Visible = 1

# Open the PowerPoint presentation
presentation = powerpoint.Presentations.Open(r"C:/Users/cnightingale/excel2slides/template_slide.pptx")

# Get a reference to the slide containing the chart
slide_index = 1  # Specify the index of the slide
slide = presentation.Slides(slide_index)
print("Retrieved slide")

# Identify the chart shape on the slide
shape = slide.Shapes(SHAPE_NAME)
print("Retrieved shape")

# Retrieve the chart object from the shape
chart = shape.Chart
print("Retrieved chart")

try:
    # Modify the chart data
    chart.ChartData.Activate()  # Activate the chart data worksheet
    
    # Access the specific range where the data is stored
    data_range = chart.ChartData.Workbook.Worksheets(1).Range("A1:B3")
    
    # Update the values in the range
    data_range.Value = [[1, "MRC"], ["DET (Cage)", 400], ["DET (Cab.)", 350]]
    
    chart.ChartData.Workbook.Close()

    # Save the presentation
    presentation.SaveAs(r"C:/Users/cnightingale/excel2slides/template_slide_modified.pptx")
    
except Exception as e:
    print("An error occurred:", str(e))

print("Chart modified and saved successfully")
# Close the presentation
presentation.Close()

# Quit the PowerPoint application
powerpoint.Quit()
print("Quit out cleanly")