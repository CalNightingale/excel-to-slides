import win32com.client
import utils
import json

class Powerpoint:
    def __init__(self, pptx_path):
        self.pptx_path = pptx_path
        self.instance = win32com.client.Dispatch("PowerPoint.Application")
        self.presentation = self.instance.Presentations.Open(pptx_path, WithWindow=False)

    def close(self):
        # Save the presentation
        self.presentation.SaveAs(r"C:/Users/cnightingale/excel2slides/template_slide_modified.pptx")
        # Close the presentation
        self.presentation.Close()
        # Quit the PowerPoint application
        self.instance.Quit()
        print("Quit out cleanly")

    def update_text(self, slide_index, element_name, new_text):
        slide = self.presentation.Slides(slide_index)
        element = slide.Shapes(element_name)
        if not element.HasTextFrame:
            raise Exception(f"Tried to set text for element '{element}' which does not have a text field")
        element.TextFrame.TextRange.Text = new_text

    def update_other(self, slide_index, element_name, data, function_name):
        try:
            function = getattr(utils, function_name)
        except:
            raise Exception(f"Failed to get function '{function_name}'... are you sure it is in utils.py?")
        slide = self.presentation.Slides(slide_index)
        element = slide.Shapes(element_name)
        function(element, data)


    def pivot_input_data(self, data):
        transposed_list = []
        # Get the length of the sublists
        sublist_length = len(data[0])
        # Iterate over the sublists, transposing the elements
        for i in range(sublist_length):
            transposed_sublist = []
            # Iterate over the main list and extract elements at index i
            for sublist in data:
                transposed_sublist.append(sublist[i])
            transposed_list.append(transposed_sublist)
        return transposed_list

    def set_chart_data(self, pptx_path : str, slide_index : int, chart_name : str, data):
        values_for_chart = self.pivot_input_data(data)
        # Get a reference to the slide containing the chart
        slide = self.presentation.Slides(slide_index)
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
            
        except Exception as e:
            print("An error occurred:", str(e))

        print("Chart modified and saved successfully")
