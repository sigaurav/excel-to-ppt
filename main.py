import win32com.client


class Generate_PPT():

    def generate_ppt(self, file_path, selected_sheets):
        ExcelApp = win32com.client.Dispatch("Excel.Application")
        ExcelApp.Visible = False
        try:
            # Open Excel workbook
            xlWorkbook = ExcelApp.Workbooks.Open(file_path)

            # Create a new instance of PowerPoint and make sure it's visible.
            PPTApp = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
            PPTApp.Visible = True

            # Add a presentation to the PowerPoint Application, returns a Presentation Object.
            PPTPresentation = PPTApp.Presentations.Add()
            # Loop through each Worksheet.
            for xlWorksheet in xlWorkbook.Worksheets:

                if xlWorksheet.Name not in selected_sheets:
                    continue
                # Grab the ChartObjects Collection for each sheet.
                xlCharts = xlWorksheet.ChartObjects()

                # Loop through each Chart in the ChartObjects Collection.
                for index, xlChart in enumerate(xlCharts):
                    # Each chart needs to be on it's own slide, so at this point create a new slide.
                    PPTSlide = PPTPresentation.Slides.Add(Index=index + 1, Layout=12)  # 12 is a blank layout

                    # Display something to the user.
                    print('Exporting Chart {} from Worksheet {}'.format(xlChart.Name, xlWorksheet.Name))

                    # Copy the chart.
                    xlChart.Copy()

                    # Paste the Object to the Slide
                    PPTSlide.Shapes.Paste()
        except Exception as e:
            print(f"Error: {e}")

        # Save the presentation.
        PPTPresentation.SaveAs(r"E:\Interview Challenges\excel-to-ppt\Output\output")