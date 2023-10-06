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

    def changed(self, name, index, mode):
        if self.var.get() == '':
            if self.lb is not None:
                self.lb.destroy()
                self.lb_up = False
        else:
            words = self.completion_list
            if words:
                if not self.lb_up:
                    self.lb = tk.Listbox()
                    self.lb.bind("<Double-Button-1>", self.selection)
                    self.lb.bind("<Right>", self.selection)
                    self.lb.place(x=self.winfo_x(), y=self.winfo_y() + self.winfo_height())
                    self.lb_up = True
                self.lb.delete(0, tk.END)
                for w in words:
                    if w.lower().startswith(self.var.get().lower()):
                        self.lb.insert(tk.END, w)
                # Resize the listbox
                self.lb.config(height=self.lb.size())
            else:
                if self.lb_up:
                    self.lb.destroy()
                    self.lb_up = False

    def selection(self, event):
        if self.lb_up:
            self.var.set(self.lb.get(tk.ACTIVE))
            self.lb.destroy()
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
        self.completion_list = sorted(completion_list, key=str.lower)

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
        
        canvas.set_image(photo_img)
        #canvas.create_image(0, 0, anchor='nw', image=photo_img)
        #canvas.image = photo_img  # Keep a reference to the image



root = tk.Tk()
root.configure(bg="white")  # Set white background
root.geometry("550x768".format(root.winfo_width()))

style = ttk.Style()
style.configure('White.TFrame', background='white')

# Create main frame
main_frame = ttk.Frame(root, padding='10', style='White.TFrame')
main_frame.pack(fill=tk.BOTH, expand=True)

# Create canvas for scrolling
canvas = tk.Canvas(main_frame)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

canvas.configure(yscrollcommand=scrollbar.set)

# Here is where canvas_frame is declared
canvas_frame = ttk.Frame(canvas, style='White.TFrame')
canvas.create_window((0, 0), window=canvas_frame, anchor=tk.NW)

canvas_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))


# Logo
logo_image = Image.open("Logo.png")
logo_image = logo_image.resize((300, int(300 * logo_image.height / logo_image.width)), Image.ANTIALIAS)  # Resize with aspect ratio
photo = ImageTk.PhotoImage(logo_image)
logo_label = tk.Label(canvas_frame, image=photo, bg="white")  # Set background
logo_label.image = photo
logo_label.pack(in_=canvas_frame, fill=tk.X, expand=True, padx=10, pady=5)

# Create a canvas to display the image
#canvas = tk.Canvas(canvas_frame, width=300, height=300, bg="white")
canvas = ImageCanvas(canvas_frame, width=300, height=300, bg="white")
canvas.pack(in_=canvas_frame, side="right")


# Create all labels, buttons, eyntry fields and dropdowns
company_entry_frame = ttk.Frame(root)
company_entry_frame.configure(style="InputFrame.TFrame")  # Set style for the entry frame
company_entry_frame.pack(in_=canvas_frame, )
company_label = tk.Label(canvas_frame, text="Company Name", bg="white", fg="black")  # Set background and text color
company_label.pack(in_=canvas_frame, fill=tk.X, expand=True, padx=10, pady=5)
company_entry = AutocompleteEntry(company_entry_frame)
company_entry.configure(background="white")  # Set background color
company_entry.pack(side="left")


date_label = tk.Label(canvas_frame, text="Date", bg="white", fg="black")  # Set background and text color
date_label.pack(in_=canvas_frame, fill=tk.X, expand=True, padx=10, pady=5)
date_entry = tk.Entry(canvas_frame, highlightcolor="#c7e9f9", highlightthickness=1)  # Set contour
date_entry.configure(background="white")  # Set background color
date_entry.pack(in_=canvas_frame, )

type_label = tk.Label(canvas_frame, text="Type", bg="white", fg="black")  # Set background and text color
type_label.pack(in_=canvas_frame, fill=tk.X, expand=True, padx=10, pady=5)
type_var = tk.StringVar(root)
type_style = ttk.Style()
type_style.configure("Type.TCombobox", background="white")  # Set background color for combobox
type_option = ttk.Combobox(canvas_frame, textvariable=type_var, values=["fonctionnement", "investissement"], state="readonly")
type_option.configure(background="white")  # Set background color
type_option.pack(in_=canvas_frame, )
type_option.bind("<<ComboboxSelected>>", update_subtype_options)  # Bind the update_subtype_options function to the selection event


subtype_label = tk.Label(canvas_frame, text="Sub-Type", bg="white", fg="black")  # Set background and text color
subtype_label.pack(in_=canvas_frame, fill=tk.X, expand=True, padx=10, pady=5)
subtype_var = tk.StringVar(root)
subtype_option = ttk.Combobox(canvas_frame, textvariable=subtype_var, values=["case1", "case2", "case3"], state="readonly")
subtype_option.configure(background="white")  # Set background color
subtype_option.pack(in_=canvas_frame, )

price_label = tk.Label(canvas_frame, text="Price (with comma, without EUR)", bg="white", fg="black")  # Set background and text color
price_label.pack(in_=canvas_frame, fill=tk.X, expand=True, padx=10, pady=5)
price_entry = tk.Entry(canvas_frame, highlightcolor="#c7e9f9", highlightthickness=1)  # Set contour
price_entry.configure(background="white")  # Set background color
price_entry.pack(in_=canvas_frame, )

input_dir = tk.StringVar(root)
output_dir = tk.StringVar(root)

input_dir_label = tk.Label(canvas_frame, text="Input Directory", bg="white", fg="black")  # Set background and text color
input_dir_label.pack(in_=canvas_frame, fill=tk.X, expand=True, padx=10, pady=5)
input_dir_button = ttk.Button(canvas_frame, text="Browse", command=browse_input_dir)
input_dir_button.pack(in_=canvas_frame, )
input_dir_entry = tk.Entry(canvas_frame, textvariable=input_dir)
input_dir_entry.pack(in_=canvas_frame, )

# Input files Listbox
input_files_label = tk.Label(canvas_frame, text="Input Files", bg="white", fg="black")  # Set background and text color
input_files_label.pack(in_=canvas_frame, fill=tk.X, expand=True, padx=10, pady=5)
input_files_listbox = tk.Listbox(canvas_frame, height=10)
input_files_listbox.pack(in_=canvas_frame, )

# Bind the listbox selection event to our function
input_files_listbox.bind('<<ListboxSelect>>', on_file_select)

output_dir_label = tk.Label(canvas_frame, text="Output Directory", bg="white", fg="black")  # Set background and text color
output_dir_label.pack(in_=canvas_frame, fill=tk.X, expand=True, padx=10, pady=5)
output_dir_button = ttk.Button(canvas_frame, text="Browse", command=browse_output_dir)
output_dir_button.pack(in_=canvas_frame, )
output_dir_entry = tk.Entry(canvas_frame, textvariable=output_dir)
output_dir_entry.pack(in_=canvas_frame, )

# Save Button
save_button = ttk.Button(canvas_frame, text="Save", command=save)
save_button.pack(in_=canvas_frame, )

# Load paths from json file if it exists
load_paths()

# Define the style for the entry frame
style = ttk.Style()
style.configure("InputFrame.TFrame", background="white", bordercolor="#c7e9f9", borderwidth=1)


root.mainloop()
