import win32com.client

# Data comes in formatted as [[A2,A3,A4,...],[B2,B3,B4,...],...]
# Need to pivot to [[A2,B2,C2,...], [A3,B3,C3,...]]
def transpose_lists(input_list):
    print(input_list)
    transposed_list = []

    # Get the length of the sublists
    sublist_length = len(input_list[0])

    # Iterate over the sublists, transposing the elements
    for i in range(sublist_length):
        transposed_sublist = []

        # Iterate over the main list and extract elements at index i
        for sublist in input_list:
            transposed_sublist.append(sublist[i])

        transposed_list.append(transposed_sublist)
    print(transposed_list)
    return transposed_list

def set_chart_data(pptx_path : str, slide_index : int, chart_name : str, data):
    values_for_chart = transpose_lists(data)

    # Create an instance of the PowerPoint application
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")

    # Open the PowerPoint presentation
    presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)

    # Get a reference to the slide containing the chart
    slide_index = 1  # Specify the index of the slide
    slide = presentation.Slides(slide_index)
    print("Retrieved slide")

    # Identify the chart shape on the slide
    shape = slide.Shapes(chart_name)
    print("Retrieved shape")

    # Retrieve the chart object from the shape
    chart = shape.Chart
    print("Retrieved chart")

    try:
        # Modify the chart data
        chart.ChartData.Activate()  # Activate the chart data worksheet

        # Access the specific range where the data is stored
        data_range = chart.ChartData.Workbook.Worksheets(1).Range("A2:B4")
        
        # Update the values in the range
        data_range.Value = values_for_chart
        
        # Close the workbook
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