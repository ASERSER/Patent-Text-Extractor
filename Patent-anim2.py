
from pdf2image import convert_from_path
import pytesseract
import os
import re
import win32com.client as win32
from PIL import Image
import tkinter as tk
from tkinter import filedialog
import os,sys
# Path to the Poppler bin directory
#poppler_path = r'C:\poppler\Library\bin'
from tqdm import tqdm
def get_tesseract_path():
    if getattr(sys, 'frozen', False):
        # If we're running as a PyInstaller bundle, use the _MEIPASS path
        return os.path.join(sys._MEIPASS, "tesseract.exe")
    else:
        # Otherwise, use the normal installed path for development/testing
        return r'C:/Program Files/Tesseract-OCR/tesseract.exe'

def show_progress(iterable, description="Processing"):
    return tqdm(iterable, desc=description, unit="step")

def select_pdf_file():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title="Select a PDF Patent File", filetypes=[("PDF Files", "*.pdf")])
    return file_path
def extract_text_from_image(image):
    """
    Extract text from a single image using pytesseract.
    """
    text = pytesseract.image_to_string(image)
    return text

def split_image(image):
    """
    Split the image into two columns.
    """
    width, height = image.size
    column_width = width // 2
    
    # Define the bounding boxes for each column
    left_bbox = (0, 0, column_width, height)
    right_bbox = (column_width, 0, width, height)
    
    # Crop the image into two columns
    left_column = image.crop(left_bbox)
    right_column = image.crop(right_bbox)
    
    return left_column, right_column

def save_image(image, output_file_path):
    """
    Save the image to a file.
    """
    image.save(output_file_path)

def convert_pdf_to_images(pdf_path, output_folder):
    """
    Convert PDF pages to images and split each image into columns.
    """
    if hasattr(sys, '_MEIPASS'):  # PyInstaller temporary folder for resources
        poppler_path = os.path.join(sys._MEIPASS, 'poppler')
    else:
        poppler_path = 'path/to/your/poppler'

    # Set DPI to improve image quality (higher DPI results in higher quality images)
    dpi_value = 600  # You can increase this to 400 or 600 for even better quality

    images = convert_from_path(pdf_path, poppler_path=poppler_path, dpi=dpi_value)
    image_paths = []
    text_data = []

    for i, image in show_progress(enumerate(images), "Converting PDF Pages to Images"):
        # Define file paths
        full_image_path = os.path.join(output_folder, f"page_{i + 1}.png")
        
        # Save the full image
        save_image(image, full_image_path)
        
        # Split the image into two columns
        left_column, right_column = split_image(image)
        
        # Define paths for column images
        left_image_path = os.path.join(output_folder, f"page_{i + 1}_col1.png")
        right_image_path = os.path.join(output_folder, f"page_{i + 1}_col2.png")
        
        # Save column images
        save_image(left_column, left_image_path)
        save_image(right_column, right_image_path)
        
        # Extract text from each column
        left_text = extract_text_from_image(left_column)
        right_text = extract_text_from_image(right_column)
        
        # Combine texts of both columns
        combined_text = left_text + "\n\n" + right_text
        text_data.append(combined_text)
        
        # Save extracted text to text file (optional)
        text_filename = f"page_{i+1}.txt"
        text_path = os.path.join(output_folder, text_filename)
        with open(text_path, 'w', encoding='utf-8') as text_file:
            text_file.write(combined_text)
        
        # Add image paths
        image_paths.append(full_image_path)
    
    return image_paths, text_data



def parse_patent_info(text):
    """
    Extract patent information based on common sections (Title, Patent Number, etc.)
    """
   
    # Use regex to match everything after (54) until another section or blank line
    title_match = re.search(r'\(54\)\s*([\s\S]+?)(?=\(\d{2}\)|\n{2,}|\Z)', text, re.IGNORECASE)
    print("The title is ",title_match)
    # If a match is found, clean up and return the title
    title = title_match.group(1).strip().replace("\n", " ") if title_match else "Title N/A"
    title = title.lstrip(')')

    patent_number_match = re.search(r'US\s\d{1,3},\d{3},\d{3}\s\w\d', text)
    patent_number = patent_number_match.group(0) if patent_number_match else "PATENT # N/A"

    date_match = re.search(r'\(45\)\s*Date of Patent:\s*(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.\s\d{1,2},\s\d{4}\b)', text)
    patent_date = date_match.group(1) if date_match else "Date N/A"

    inventors_match = re.search(r'Inventors?:\s*([\s\S]+?)(?=\(\d{2}\)|\n\n|\Z)', text, re.IGNORECASE)
    inventors = inventors_match.group(1).strip().replace("\n", " ") if inventors_match else "Inventors N/A"

    abstract_match = re.search(r'Abstract:?\s*([\s\S]+?)(?=\n\n|\Z)', text, re.IGNORECASE)
    abstract = abstract_match.group(1).strip() if abstract_match else "Abstract N/A"

    return title, patent_number, patent_date, inventors, abstract
def modify_active_ppt(images, text_data):
    """
    Modify the active PowerPoint presentation with images and extracted text in a clean and elegant layout.
    """
    # Open the active PowerPoint application
    powerpoint = win32.Dispatch("PowerPoint.Application")
    
    # Reference the active presentation (already open)
    presentation = powerpoint.ActivePresentation

    # Delete existing slides (if needed, can be customized)
    for slide_index in range(presentation.Slides.Count, 1, -1):
        presentation.Slides(slide_index).Delete()

    # Loop through each image and corresponding text data
    for i, (image_path, text) in show_progress(enumerate(zip(images, text_data)), "Modifying PowerPoint Slides"):
        # Add a new slide based on the layout of the first slide in the active presentation
        blank_layout = 12  # Assuming layout type 12 for blank
        slide = presentation.Slides.Add(presentation.Slides.Count + 1, blank_layout)

        # Set a sleek background color (dark matte look)
        slide.FollowMasterBackground = False
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = 0x1C1C1C  # Dark gray color
        
        # Extract patent information
        title, patent_number, patent_date, inventors, abstract = parse_patent_info(text)

        # Insert the patent image on the right side of the slide
        picture_shape =slide.Shapes.AddPicture(
            image_path, LinkToFile=False, SaveWithDocument=True, 
            Left=500, Top=50, Width=374.4, Height=459.6
        )
        picture_shape.name = "patent_image"

        # Set a specific style for text boxes with clean and elegant formatting
        title_shape = slide.Shapes.AddTextbox(Orientation=1, Left=50, Top=40, Width=400, Height=50)
        title_shape.TextFrame.TextRange.Text = f"{title}"
        title_shape.TextFrame.TextRange.Font.Bold = True
        title_shape.TextFrame.TextRange.Font.Size = 16
        title_shape.TextFrame.TextRange.Font.Name = "Calibri"
        title_shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF  # White text for contrast

        patent_shape = slide.Shapes.AddTextbox(Orientation=1, Left=50, Top=120, Width=400, Height=30)
        patent_shape.TextFrame.TextRange.Text = f"PATENT #: {patent_number}     {patent_date}"
        patent_shape.TextFrame.TextRange.Font.Size = 16
        patent_shape.TextFrame.TextRange.Font.Name = "Calibri"
        patent_shape.TextFrame.TextRange.Font.Color.RGB = 0xD3D3D3  # Light gray for subtle emphasis
        text_frame = patent_shape.TextFrame
        text_range = text_frame.TextRange

        # Find the position of the patent number and apply bold formatting to it
        start_pos = text_range.Text.find(patent_number)
        if start_pos != -1:
            end_pos = start_pos + len(patent_number)
            text_range.Characters(start_pos + 1, len(patent_number)).Font.Bold = True

             # Create a text box in PowerPoint
        inventors_shape = slide.Shapes.AddTextbox(Orientation=1, Left=50, Top=160, Width=400, Height=60)
        inventors_shape.TextFrame.TextRange.Font.Size = 14
        inventors_shape.TextFrame.TextRange.Font.Name = "Calibri"
        inventors_shape.TextFrame.TextRange.Font.Color.RGB = 0xD3D3D3
        inventors_shape.TextFrame.TextRange.Text = f"INVENTORS: {inventors}"
        inventors_shape.name = "InventorsTextBox" 
                # Bold the inventors' names (the part after "INVENTORS: ")
        inventors_start_pos = len("INVENTORS: ") + 1  # Position where the inventors' names start
        inventors_shape.TextFrame.TextRange.Characters(inventors_start_pos, len(inventors)).Font.Bold = True
        

        # Adjust the position of the abstract depending on the length of the inventors' list
        inventors_height = inventors_shape.Height
        abstract_top_position = 160 + inventors_height + 10  # Adjust top position dynamically based on inventors' box height

        # Create an abstract text box
        abstract_shape = slide.Shapes.AddTextbox(Orientation=1, Left=50, Top=abstract_top_position, Width=400, Height=200)
        
        # Truncate the abstract to fit within the text box
        abstract_text = f"ABSTRACT:\n{abstract}"
        abstract_shape.TextFrame.TextRange.Text = abstract_text
        abstract_shape.TextFrame.TextRange.Font.Size = 14
        abstract_shape.TextFrame.TextRange.Font.Name = "Calibri"
        abstract_shape.TextFrame.TextRange.Font.Color.RGB = 0xB0B0B0  # Slightly lighter gray
        abstract_shape.TextFrame.AutoSize = 1  # Auto-sizing enabled for abstract as well
        abstract_shape.name ="abstractbox"
        # Reduce font size if text overflows
        max_height = 200  # Max height allowed for abstract text box
        while abstract_shape.Height > max_height and abstract_shape.TextFrame.TextRange.Font.Size > 8:
            # Reduce font size
            abstract_shape.TextFrame.TextRange.Font.Size -= 1

        # Optional thin line separator for added elegance
        line = slide.Shapes.AddLine(BeginX=50, BeginY=abstract_top_position - 10, EndX=450, EndY=abstract_top_position - 10)
        line.Line.ForeColor.RGB = 0x888888  # Medium gray line
        line.Line.Weight = 0.75
        line.name="patent_line"
    # Save the updated presentation
    presentation.Save()
    print("PowerPoint presentation updated.")



def main():
    import shutil
    
    output_folder = r"C:\Images"

    # Check if the directory exists
    if os.path.exists(output_folder):
        # If it exists, delete the directory and all its contents
        shutil.rmtree(output_folder)
        print(f"Directory '{output_folder}' removed.")
    else:
        print(f"Directory '{output_folder}' does not exist.")

    # Create the directory again
    os.makedirs(output_folder)
    print(f"Directory '{output_folder}' created.")
    pdf_path = select_pdf_file()
    
    if not pdf_path:
        print("No file selected. Exiting.")
        return
    

    # Convert the PDF to images and extract text
    image_paths, text_data = convert_pdf_to_images(pdf_path, output_folder)

    # Modify the active PowerPoint presentation
    modify_active_ppt(image_paths, text_data)

if __name__ == "__main__":
    main()
