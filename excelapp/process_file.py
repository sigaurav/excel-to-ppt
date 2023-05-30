import win32com.client
import os
from django.http import JsonResponse, HttpResponse
from django.conf import settings

def Generate_ppt(file_path, selected_sheets):

    if not selected_sheets:
        return HttpResponse("No sheets selected")

    # Update the file path based on the uploaded Excel file
    output_path = os.path.join(settings.MEDIA_ROOT, 'outputFiles', 'output.pptx')

    try:
        ExcelApp = win32com.client.Dispatch("Excel.Application")
        ExcelApp.Visible = False

        # Open the Excel workbook
        xlWorkbook = ExcelApp.Workbooks.Open(file_path)

        # Create a new instance of PowerPoint and make sure it's visible.
        PPTApp = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
        PPTApp.Visible = False

        # Add a presentation to the PowerPoint Application, returns a Presentation Object.
        PPTPresentation = PPTApp.Presentations.Add()

        # Loop through each worksheet in the Excel workbook
        for xlWorksheet in xlWorkbook.Worksheets:
            if xlWorksheet.Name not in selected_sheets:
                continue

            # Grab the ChartObjects Collection for each sheet.
            xlCharts = xlWorksheet.ChartObjects()

            # Loop through each Chart in the ChartObjects Collection.
            for index, xlChart in enumerate(xlCharts):
                # Each chart needs to be on its own slide, so create a new slide.
                PPTSlide = PPTPresentation.Slides.Add(Index=index + 1, Layout=12)  # 12 is a blank layout

                # Copy the chart.
                xlChart.Copy()

                # Paste the Object to the Slide
                PPTSlide.Shapes.Paste()

        # Save the presentation
        PPTPresentation.SaveAs(output_path)

        return JsonResponse({'file_name': 'output.pptx'})

    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)

