import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
import docx2pdf

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Word to PDF Converter")

        # Create a button to select the Word file
        self.browse_button = tk.Button(self.root, text="Browse", command=self.browse_file)
        self.browse_button.pack()

        # Create a button to convert the Word file to PDF
        self.convert_button = tk.Button(self.root, text="Convert", command=self.convert_file)
        self.convert_button.pack()

    def browse_file(self):
        # Open a file dialog to select the Word file
        self.filepath = filedialog.askopenfilename(title="Select Word file", filetypes=[("Word files", "*.docx")])

    def convert_file(self):
        # Convert the Word file to PDF
        try:
            pdf_file = docx2pdf.convert(self.filepath)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        # Save the PDF file
        filedialog.asksaveasfile(mode="wb", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialfile="converted.pdf")


root = tk.Tk()
app = App(root)
root.mainloop()

