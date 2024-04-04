import os
import sys
import ctypes
from pypdf import PdfWriter, PdfReader
import tkinter as tk
from tkinter import filedialog

if getattr(sys, 'frozen', False):
    import pyi_splash

class GUI(tk.Tk):
    def __init__(self):
        super().__init__()

        window_width = 470
        window_height = 200

        self.title("PDF Page Extractor")
        self.geometry(f"{window_width}x{window_height}")

        # Calculate the center position
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        position_top = int(screen_height / 2 - window_height / 2)
        position_right = int(screen_width / 2 - window_width / 2)

        # Position the window
        self.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

        self.input_file = tk.StringVar()
        self.pages = tk.StringVar()

        self.input_file_label = tk.Label(self, text="Input File")
        self.input_file_label.pack()

        self.input_file_entry = tk.Entry(self, textvariable=self.input_file, width=70)
        self.input_file_entry.pack()

        self.browse_button = tk.Button(self, text="Browse", command=self.browse, width=15)
        self.browse_button.pack(pady=(5, 10))

        self.pages_label = tk.Label(self, text="Pages (e.g. 1,3-6,9)")
        self.pages_label.pack()

        self.pages_entry = tk.Entry(self, textvariable=self.pages, width=70)
        self.pages_entry.pack()

        self.extract_button = tk.Button(self, text="Extract", command=self.extract, width=15)
        self.extract_button.pack(pady=(5))

        # Create the label with an empty text
        self.extract_complete_label = tk.Label(self, text="", fg="green")
        self.extract_complete_label.pack(pady=(2))

        # Add clickable link to open github page
        self.github_link = tk.Label(self, text="github.com/jmclaren7/pdf-page-extractor", fg="blue", cursor="hand2")
        self.github_link.place(x=window_width, y=window_height, anchor="se")
        self.github_link.bind("<Button-1>", lambda e: os.system("start https://github.com/jmclaren7/pdf-page-extractor"))

        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # Hide the console window
        if getattr(sys, 'frozen', False):
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

    def browse(self):
        # Open a file dialog to select a PDF file
        self.input_file.set(filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf"), ("Any", "*.*")]))

    def extract(self):
        # Update the label to show that extraction is in progress
        self.extract_complete_label.config(text="...", fg="black")
        
        # Get the input file and pages from the GUI
        input_file = self.input_file.get()
        pages = self.pages.get()

        # Check if input file and pages are provided
        if not input_file or not pages:
            return

        # Generate the output file name
        output_file = os.path.splitext(input_file)[0] + "_extracted.pdf"
        
        try:
            # Call the extract_pages function with the provided arguments
            result = extract_pages(input_file, pages, output_file)
        except:
            # Handle any unknown errors
            self.extract_complete_label.config(text="Error: unknown", fg="red")
        else:
            if result == -1:
                # Handle the case when the page range exceeds the total number of pages
                self.extract_complete_label.config(text="Error: page range exceeded", fg="red")
            else:
                # Update the label to show the number of extracted pages and the output file name
                filename = os.path.basename(output_file)
                self.extract_complete_label.config(text="Extracted {} pages to: {}".format(result, filename), fg="green")

    def on_close(self):
        # Terminate the application
        self.destroy()
        sys.exit(0)


def extract_pages(input_file, pages, output_file):
    # Split the pages string into a list of page numbers and split ranges into individual pages
    pdf_pages = []
    for page in pages.split(','):
        if '-' in page:
            start, end = map(int, page.split('-'))
            pdf_pages.extend(range(start, end+1))
        else:
            pdf_pages.append(int(page))

    # Create a PdfWriter object to store the extracted pages
    output = PdfWriter()

    # Read the input PDF file
    input_pdf = PdfReader(input_file)

    # Check if the specified pages are within the range of the input PDF
    num_pages = len(input_pdf.pages)
    for page in pdf_pages:
        try:
            assert page <= num_pages
        except:
            return -1

    # Extract the specified pages and add them to the output PdfWriter object
    for i in pdf_pages:
        output.add_page(input_pdf.pages[i-1])

    # Write the extracted pages to the output file
    with open(output_file, 'wb') as output_stream:
        output.write(output_stream)

    # Return the number of extracted pages
    return len(output.pages)



if __name__ == '__main__':
    if getattr(sys, 'frozen', False):
        pyi_splash.close()
    gui = GUI()
    gui.mainloop()
