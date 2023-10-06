import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os
from datetime import datetime
from PIL import Image, ImageTk
import shutil
import fitz

# Function to save paths
def save_paths():
    paths = {"input_dir": input_dir.get(), "output_dir": output_dir.get()}
    with open('paths.json', 'w') as f:
        json.dump(paths, f)

# Function to load paths
def load_paths():
    if os.path.exists('paths.json'):
        with open('paths.json', 'r') as f:
            paths = json.load(f)
        input_dir.set(paths.get('input_dir', ''))
        output_dir.set(paths.get('output_dir', ''))

# Function to browse input directory
def browse_input_dir():
    dirname = filedialog.askdirectory()
    input_dir.set(dirname)
    save_paths()
    update_input_files_listbox()

# Function to browse output directory
def browse_output_dir():
    dirname = filedialog.askdirectory()
    output_dir.set(dirname)
    save_paths()

# Function to update the options in the subtype menu when the type is changed
def update_subtype_options(*args):
    selected_type = type_option.get()
    if selected_type == "fonctionnement":
        subtype_option['values'] = ["case1", "case2", "case3"]
        with open('fonctionnement_list.txt', 'r') as f:
            companies = f.read().splitlines()
        company_entry.set_completion_list(companies)
    else:
        subtype_option['values'] = ["Travaux", "Matériel", "Pédagogie", "Jeux", "electromenager"]
        with open('investissement_list.txt', 'r') as f:
            companies = f.read().splitlines()
        company_entry.set_completion_list(companies)
    subtype_option.current(0)  # Select the first option by default


# Save function
def save():
    # Check that all fields are filled
    if not all([input_dir.get(), output_dir.get(), company_entry.get(), date_entry.get(), type_var.get(), price_entry.get()]):
        messagebox.showerror("Error", "Please fill all fields")
        return
    # Check that the selected date is valid
    try:
        datetime.strptime(date_entry.get(), "%Y%m%d")
    except ValueError:
        messagebox.showerror("Error", "Invalid date. Please enter date in YYYYMMDD format")
        return
    # Check if the output directory exists, if not, create it
    output_directory = output_dir.get()
    if not os.path.isdir(output_directory):
        os.makedirs(output_directory)
    # Check if the type subdirectory exists, if not, create it
    if not os.path.isdir(os.path.join(output_directory, type_var.get())):
        os.makedirs(os.path.join(output_directory, type_var.get()))
    # Get the selected file from the listbox
    selected_file = input_files_listbox.get(tk.ACTIVE)
    if not selected_file:
        messagebox.showerror("Error", "No file selected")
        return
    # Define the new filename and filepath
    new_filename = "{} - {} - {} {}EUR.pdf".format(date_entry.get(), company_entry.get(), subtype_var.get(), price_entry.get())
    new_filepath = os.path.join(output_directory, type_var.get(), new_filename)
    # Move and rename the selected file
    shutil.move(os.path.join(input_dir.get(), selected_file), new_filepath)
    # Update the list of input files
    update_input_files_listbox()

class AutocompleteEntry(ttk.Entry):
    def __init__(self, *args, **kwargs):
        ttk.Entry.__init__(self, *args, **kwargs)
        self.var = self["textvariable"]
        if self.var == '':
            self.var = self["textvariable"] = tk.StringVar()
        self.completion_list = []
        self.var.trace('w', self.changed)
        self.bind("<Right>", self.selection)
        self.bind("<Up>", self.up)
        self.bind("<Down>", self.down)
        self.lb_up = False
        self.lb = None
        self.toplevel = None

    def changed(self, name, index, mode):
        if self.var.get() == '':
            if self.toplevel:
                self.toplevel.destroy()
                self.lb_up = False
        else:
            words = self.completion_list
            if words:
                if not self.lb_up:
                    self.toplevel = tk.Toplevel(self)
                    self.toplevel.wm_overrideredirect(True)
                    self.toplevel.geometry('+%d+%d' % (self.winfo_rootx(), self.winfo_rooty() + self.winfo_height()))
                    self.lb = tk.Listbox(self.toplevel)
                    self.lb.bind("<Double-Button-1>", self.selection)
                    self.lb.bind("<Right>", self.selection)
                    self.lb.pack()
                    self.lb_up = True
                self.lb.delete(0, tk.END)
                for w in words:
                    if w.lower().startswith(self.var.get().lower()):
                        self.lb.insert(tk.END, w)
            else:
                if self.lb_up:
                    self.toplevel.destroy()
                    self.lb_up = False

    def selection(self, event):
        if self.lb_up:
            self.var.set(self.lb.get(tk.ACTIVE))
            self.toplevel.destroy()
            self.lb_up = False
            self.icursor(tk.END)

    def up(self, event):
        if self.lb_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = self.lb.curselection()[0]
            if index != '0':
                self.lb.selection_clear(first=index)
                index = str(int(index) - 1)
                self.lb.selection_set(first=index)
                self.lb.activate(index)

    def down(self, event):
        if self.lb_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = self.lb.curselection()[0]
            self.lb.selection_clear(first=index)
            index = str(int(index) + 1)
            self.lb.selection_set(first=index)
            self.lb.activate(index)

    def set_completion_list(self, completion_list):
        self.completion_list = sorted(list(set(completion_list)), key=str.lower)

class ImageCanvas(tk.Canvas):
    def __init__(self, parent, *args, **kwargs):
        tk.Canvas.__init__(self, parent, *args, **kwargs)
        self.image_id = None
        self.image = None
        self.bind("<ButtonPress-1>", self.drag_start)
        self.bind("<B1-Motion>", self.drag_move)
        self.bind("<MouseWheel>", self.zoom)

    def set_image(self, image):
        self.image = image
        if self.image_id is not None:
            self.delete(self.image_id)
        self.image_id = self.create_image(0, 0, anchor='nw', image=image)

    def drag_start(self, event):
        self.scan_mark(event.x, event.y)

    def drag_move(self, event):
        self.scan_dragto(event.x, event.y, gain=1)

    def zoom(self, event):
        scale = 1.0
        # Respond to Linux (event.num) or Windows (event.delta) wheel event
        if event.num == 5 or event.delta < 0:
            scale -= 0.1
        if event.num == 4 or event.delta > 0:
            scale += 0.1
        if self.image:
            x = self.winfo_pointerx() - self.winfo_rootx()
            y = self.winfo_pointery() - self.winfo_rooty()
            bbox = self.bbox(self.image_id)
            self.scale(self.image_id, x, y, scale, scale)
            self.configure(scrollregion=self.bbox("all"))

def update_input_files_listbox():
    # Clear the listbox
    input_files_listbox.delete(0, tk.END)

    # List all PDF and image files in the input directory
    input_directory = input_dir.get()
    for filename in os.listdir(input_directory):
        if filename.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg')):
            if is_duplicate(filename):
                input_files_listbox.insert(tk.END, filename)
                input_files_listbox.itemconfig(tk.END, {'fg': 'red'})
            else:
                input_files_listbox.insert(tk.END, filename)

# Function to check if a file in the input directory has the same size as any file in the output directory or its subdirectories
def is_duplicate(file):
    input_file_size = os.path.getsize(os.path.join(input_dir.get(), file))
    for root, dirs, files in os.walk(output_dir.get()):
        for filename in files:
            output_file_size = os.path.getsize(os.path.join(root, filename))
            if input_file_size == output_file_size:
                return True
    return False

# Function to convert the first page of PDF to an image
def pdf_to_img(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)  # number of page
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img


def on_file_select(event):
    # Get your input directory from the label
    in_folder = input_dir_entry.get()

    # Get selected line index
    index = input_files_listbox.curselection()
    
    # Get the listbox's list of files
    if index:  # If there is any selection
        filename = input_files_listbox.get(index[0])
        # Now you have the filename and can do whatever you want with it
        pdf_path = os.path.join(in_folder, filename)  # Complete path to the pdf file
        img = pdf_to_img(pdf_path)
        # img.save('debug_output.png')  # Save the output for debugging
        photo_img = ImageTk.PhotoImage(img)

        # Set the canvas width and height to the dimensions of the image
        canvas.config(width=photo_img.width(), height=photo_img.height())
        
        pdf_canvas.set_image(photo_img)
        #canvas.create_image(0, 0, anchor='nw', image=photo_img)
        #canvas.image = photo_img  # Keep a reference to the image


root = tk.Tk()
root.configure(bg="white")
root.geometry("800x768")

# Define the style for the entry frame
style = ttk.Style()
style.configure('White.TFrame', background="white", bordercolor="#c7e9f9", borderwidth=1)

# Create main frame
main_frame = ttk.Frame(root, padding='10', style='White.TFrame')
main_frame.pack(fill=tk.BOTH, expand=True)

# Create canvas for scrolling
canvas = tk.Canvas(main_frame, bg="white")
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

canvas.configure(yscrollcommand=scrollbar.set)

# Frame inside the canvas
canvas_frame = ttk.Frame(canvas, style='White.TFrame')
canvas.create_window((0, 0), window=canvas_frame, anchor=tk.NW)

canvas_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

# Logo
logo_image = Image.open("Logo.png")
logo_image = logo_image.resize((300, int(300 * logo_image.height / logo_image.width)), Image.ANTIALIAS)
photo = ImageTk.PhotoImage(logo_image)
logo_label = tk.Label(canvas_frame, image=photo, bg="white")
logo_label.image = photo
logo_label.grid(row=0, column=0, padx=10, pady=5)

# Create a separate frame for the ImageCanvas
canvas_container = ttk.Frame(canvas_frame)
canvas_container.grid(row=0, column=2, rowspan=8, sticky="nsew")

# Create a canvas to display the image
pdf_canvas = ImageCanvas(canvas_container, width=300, height=300, bg="white")  # renamed to avoid conflict
pdf_canvas.pack(fill="both", expand=True)

# Create a frame for the central column
center_frame = ttk.Frame(canvas_frame, style='White.TFrame')
center_frame.grid(row=1, column=0, rowspan=8, sticky="nsew")  # spans 8 rows

# Adjust other widgets inside center_frame
company_label = tk.Label(center_frame, text="Company Name", bg="white", fg="black")

company_entry_frame = ttk.Frame(center_frame, style="InputFrame.TFrame")

company_entry = AutocompleteEntry(company_entry_frame)
company_entry.configure(background="white")
company_entry.pack(side="left")

date_label = tk.Label(center_frame, text="Date", bg="white", fg="black")

date_entry = tk.Entry(center_frame, highlightcolor="#c7e9f9", highlightthickness=1)
date_entry.configure(background="white")

type_label = tk.Label(center_frame, text="Type", bg="white", fg="black")

type_var = tk.StringVar(root)
type_style = ttk.Style()
type_style.configure("Type.TCombobox", background="white")
type_option = ttk.Combobox(center_frame, textvariable=type_var, values=["fonctionnement", "investissement"], state="readonly")
type_option.configure(background="white")
type_option.bind("<<ComboboxSelected>>", update_subtype_options)  # Added binding here

subtype_label = tk.Label(center_frame, text="Sub-Type", bg="white", fg="black")

subtype_var = tk.StringVar(root)
subtype_option = ttk.Combobox(center_frame, textvariable=subtype_var, values=["case1", "case2", "case3"], state="readonly")
subtype_option.configure(background="white")

price_label = tk.Label(center_frame, text="Price (with comma, without EUR)", bg="white", fg="black")

price_entry = tk.Entry(center_frame, highlightcolor="#c7e9f9", highlightthickness=1)
price_entry.configure(background="white")

input_dir = tk.StringVar(root)
output_dir = tk.StringVar(root)

input_dir_label = tk.Label(center_frame, text="Input Directory", bg="white", fg="black")

input_dir_button = ttk.Button(center_frame, text="Browse", command=browse_input_dir)

input_dir_entry = tk.Entry(center_frame, textvariable=input_dir)

# Input files Listbox
input_files_label = tk.Label(center_frame, text="Input Files", bg="white", fg="black")

input_files_listbox = tk.Listbox(center_frame, height=10)
input_files_listbox.bind('<<ListboxSelect>>', on_file_select)

output_dir_label = tk.Label(center_frame, text="Output Directory", bg="white", fg="black")

output_dir_button = ttk.Button(center_frame, text="Browse", command=browse_output_dir)

output_dir_entry = tk.Entry(center_frame, textvariable=output_dir)

# Save Button
save_button = ttk.Button(center_frame, text="Save", command=save)

center_frame.grid(row=1, column=0, rowspan=8, sticky="nsew")  # spans 8 rows

# Adjust other widgets inside center_frame (From row 1 to row 8)
company_label.grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
company_entry_frame.grid(row=2, column=0, padx=10, pady=5)
date_label.grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
date_entry.grid(row=4, column=0, padx=10, pady=5)
type_label.grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
type_option.grid(row=6, column=0, padx=10, pady=5)
subtype_label.grid(row=7, column=0, sticky=tk.W, padx=10, pady=5)
subtype_option.grid(row=8, column=0, padx=10, pady=5)

# Place the remaining widgets in column 1 (From row 1 to row 19)
price_label.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
price_entry.grid(row=2, column=1, padx=10, pady=5)
input_dir_label.grid(row=3, column=1, sticky=tk.W, padx=10, pady=5)
input_dir_button.grid(row=4, column=1, sticky=tk.W, padx=10, pady=5)
input_dir_entry.grid(row=5, column=1, padx=10, pady=5)
input_files_label.grid(row=6, column=1, sticky=tk.W, padx=10, pady=5)
input_files_listbox.grid(row=7, column=1, padx=10, pady=5)
output_dir_label.grid(row=8, column=1, sticky=tk.W, padx=10, pady=5)
output_dir_button.grid(row=9, column=1, sticky=tk.W, padx=10, pady=5)
output_dir_entry.grid(row=10, column=1, padx=10, pady=5)
save_button.grid(row=11, column=1, padx=10, pady=5)

# Load paths from json file if it exists
load_paths()

root.mainloop()
