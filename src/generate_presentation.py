import win32com.client
import os
import sys

def create_presentation():
    ppt = None
    presentation = None
    try:
        # Paths (update if needed)
        template_path = os.path.abspath('template.pptx')
        plot_path = os.path.abspath('results/income_plot.png')
        output_pdf = os.path.abspath('presentation.pdf')

        # Validate files
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Missing template: {template_path}")
        if not os.path.exists(plot_path):
            raise FileNotFoundError(f"Missing plot: {plot_path}")

        # Launch PowerPoint
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.Visible = True

        # Open template
        presentation = ppt.Presentations.Open(template_path)

        # Slide 1: Title
        slide1 = presentation.Slides.Add(1, 1)
        slide1.Shapes.Title.TextFrame.TextRange.Text = "Automated Report"
        slide1.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Generated with Python"

        # Slide 2: Plot
        slide2 = presentation.Slides.Add(2, 2)
        slide2.Shapes.Title.TextFrame.TextRange.Text = "Income Distribution"
        slide2.Shapes.AddPicture(plot_path, 0, 1, 100, 100, Width=500, Height=300)

        # Slide 3: Model Results
        slide3 = presentation.Slides.Add(3, 2)  # Title and Content layout
        slide3.Shapes.Title.TextFrame.TextRange.Text = "Predictive Model Performance"

        # Read accuracy from file
        with open('results/model_metrics.txt', 'r') as f:
            accuracy = f.read()

        # Add text box
        left = top = 100  # Position
        width, height = 500, 50
        textbox = slide3.Shapes.AddTextbox(1, left, top, width, height)
        textbox.TextFrame.TextRange.Text = accuracy

        presentation.SaveAs(output_pdf, FileFormat=32)

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

    finally:
        if presentation:
            presentation.Close()
        if ppt:
            ppt.Quit()

if __name__ == "__main__":
    create_presentation()