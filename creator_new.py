import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from fpdf import FPDF
from pdf2image import convert_from_bytes
from PIL import ImageTk, Image
import os
import platform
import csv
from datetime import datetime

global folder_name
folder_name ="Wohnung-Gardasee"

def load_data_from_csv(file_path, lang="de"):
    data = {}
    with open(file_path, newline="", encoding="utf-8") as csvfile:
        reader = csv.reader(csvfile)
        if lang=="de":
            for row in reader:
                key = row[0]
                data[key] = row[1].strip()
                
        if lang=="it":
            for row in reader:
                key = row[0]
                data[key] = row[2].strip()
                
        if lang=="en":
            for row in reader:
                key = row[0]
                data[key] = row[3].strip()
                
    return data

pages = {
    "KEY_DATA":              999,
    "DESCRIPTION":           999,
    "IMPRESSIONS":           999,
    "LOCATION_DESCRIPTION":  999,
    "FLOOR_PLANS":           999,
    "AREA_CALCULATION":      999,
    "PROVISION":             999,
    "CONTACT_PERSON":        999,
}
global lang 
lang ="de"
global text_list 
text_list = load_data_from_csv("Allgemein/DatenAllgemein.csv")

PAGE_WIDTH = 210
PAGE_HEIGHT = 297
IMAGE_HEIGHT = PAGE_HEIGHT * 0.43
Y_POSITIONS = [20, 150]

class ExposePDF(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=5)
        self.add_font(
            "DejaVu", "", "./Barlow-fontiko/Barlow Regular.ttf", uni=True
        )  # <-- Unicode font
        self.add_font(
            "DejaVu", "I", "./Barlow-fontiko/Barlow Italic.ttf", uni=True
        )  # Italic font
        self.add_font(
            "DejaVu", "B", "./Barlow-fontiko/Barlow Bold.ttf", uni=True
        )  # Italic font
        
        self.add_font(
            "DejaVu-TITLE", "", "./Barlow-fontiko/Barlow Medium.ttf", uni=True
        )  # Italic font
        self.set_font("DejaVu", "", 14)

    def header(self):
        pass  # No default header

    def footer(self):
        if (self.page_no()>2):
           self.set_y(-15)  # Position 15mm from bottom
           self.set_font("DejaVu", "I", 8)

           # Left-aligned "PREMIUM HOMES"
           self.cell(0, 10, text_list["SUBTITLE"], align="L")

           # Page number right-aligned
           self.set_x(-30)
#           self.cell(0, 10, f"{text_list["PAGE"]} {self.page_no()-2}", align="R")


def list_single(pdf, left, line=False):
    pdf.set_font("DejaVu", "", 11)

    start_x_left = 10
    start_y = pdf.get_y()

    line_height = 10
    col_width = 190

    for i in range(len(left)):
        y = start_y + i * line_height

        if i < len(left):
            pdf.set_xy(start_x_left, y)
            pdf.set_font("DejaVu", "", 10)
            pdf.cell(
                60, line_height, left[i][0]
            )  # Left column, first part (no alignment change)
            pdf.set_font("DejaVu", "", 10)
            # Align left[i][1] to the right
            pdf.cell(130, line_height, left[i][1], align="R")

            # draw line under left row
            if line:
                pdf.line(
                    start_x_left,
                    y + line_height-3,
                    start_x_left + col_width,
                    y + line_height-3,
                )


def list_double(pdf, left, right, line=False):

    pdf.set_font("DejaVu", "", 11)

    start_x_left = 10
    start_x_right = 110
    start_y = pdf.get_y()

    line_height = 10
    col_width = 90

    for i in range(max(len(left), len(right))):
        y = start_y + i * line_height

        if i < len(left):
            pdf.set_xy(start_x_left, y)
            pdf.set_font("DejaVu", "", 10)
            pdf.cell(60, line_height, left[i][0]) 
            pdf.set_font("DejaVu", "", 10)
            pdf.cell(30, line_height, left[i][1], align="R")
            if line:
                pdf.line(
                    start_x_left,
                    y + line_height - 3,
                    start_x_left + col_width,
                    y + line_height - 3,
                )

        if i < len(right):
            pdf.set_xy(start_x_right, y)
            pdf.set_font("DejaVu", "", 10)
            pdf.cell(60, line_height, right[i][0])
            pdf.set_font("DejaVu", "", 10)
            pdf.cell(30, line_height, right[i][1], align="R")
            if line:
                pdf.line(
                    start_x_right,
                    y + line_height -3,
                    start_x_right + col_width,
                    y + line_height -3,
                )


def page_title(pdf, csv_data, title):
    pdf.set_font("DejaVu-TITLE", "", 18)
    pdf.cell(0, 10, title, ln=True, align="C")
    pdf.ln(5)
    pdf.set_font("DejaVu", "", 16)
    pdf.cell(0, 10, csv_data["UNTERTITEL"], ln=True, align="C")
    pdf.ln(15)
    pdf.set_font("DejaVu", "", 11)


def split_list(data_list):
    midpoint = (len(data_list) + 1) // 2  
    left = data_list[:midpoint]
    right = data_list[midpoint:]
    return left, right

def valid():
    return datetime.now() < datetime(2025, 9, 1)

def create_pdf():
    
    pdf = ExposePDF()
    pdf.footer()
    csv_data = load_data_from_csv(folder_name+"/Daten.csv",lang)

    # First Page
    pdf.add_page()
    pdf.set_font("DejaVu-TITLE", "", 16)
    pdf.ln(1)
    pdf.image("Allgemein/logo.jpeg", x=(210 - 60) / 2, w=60)
    pdf.ln(5)
    pdf.image(folder_name+"/"+csv_data["TITEL_BILD"], x=(210 - 190) / 2, w=190)
    pdf.ln(2)
    pdf.multi_cell(
        0,
        8,
        csv_data["TITEL"],
        align="C",
    )
    pdf.ln(2)
    pdf.set_font("DejaVu", "", 16)
    pdf.cell(0, 10, csv_data["UNTERTITEL"], ln=True, align="C")
    pdf.ln(3)
    pdf.set_font("DejaVu", "", 11)
    
    list = [
        (text_list["TYPE"]                  , csv_data["PROPERTY_TYPE"]),
        (text_list["ROOM_FOR_PAYMENT"]      , csv_data["ROOM_FOR_PAYMENT"]),
        (text_list["RETAIL_SPACE"]          , csv_data["RETAIL_SPACE"]),
        (text_list["PURCHASE_PRICE"]        , csv_data["PURCHASE_PRICE"]),
    ]
    list = [(key, value) for key, value in list if value]
    left, right = split_list(list)
    
    list_double(pdf, left=left, right=right, line=True)
    pdf.ln(40)
    pdf.set_font("DejaVu", "B", 11)
    pdf.set_y(-25)
    pdf.cell(0, 5,text_list["ESTATE_AGENT"], ln=True, align="L")

    pdf.set_font("DejaVu", "", 11)
    pdf.cell(0, 5,text_list["EMAIL"] , ln=False, align="L")
    pdf.cell(0, 5,text_list["WEBSITE"] , ln=True, align="R")
    pdf.cell(0, 5,text_list["TELEPHONE"] , ln=False, align="L")
    pdf.cell(0, 5,text_list["COMPANY"] , ln=True, align="R")

    # Table of Contents
    pdf.add_page()
    page_title(pdf=pdf, csv_data=csv_data, title=text_list["OVERVIEW"].upper())

    left = [
        (text_list["KEY_DATA"],              text_list["PAGE"] + pages["KEY_DATA"].__str__()),
        (text_list["DESCRIPTION"],           text_list["PAGE"] + pages["DESCRIPTION"].__str__()),
        (text_list["IMPRESSIONS"],           text_list["PAGE"] + pages["IMPRESSIONS"].__str__()),
        (text_list["LOCATION_DESCRIPTION"],  text_list["PAGE"] + pages["LOCATION_DESCRIPTION"].__str__()),
        (text_list["FLOOR_PLANS"],           text_list["PAGE"] + pages["FLOOR_PLANS"].__str__()),
        (text_list["AREA_CALCULATION"],      text_list["PAGE"] + pages["AREA_CALCULATION"].__str__()),
        (text_list["PROVISION"],             text_list["PAGE"] + pages["PROVISION"].__str__()),
        (text_list["CONTACT_PERSON"],        text_list["PAGE"] + pages["CONTACT_PERSON"].__str__()),
    ]

    list_single(pdf, left=left, line=True)

    pdf.set_y(-150)

    image_files = sorted(
        [
            os.path.join(folder_name, f)
            for f in os.listdir(folder_name)
            if f.startswith("ÜBERSICHT")
            and f.lower().endswith((".jpg", ".jpeg", ".png"))
        ]
    )
    img_path = image_files[0]
    if os.path.exists(img_path):
        with Image.open(img_path) as img:
            x = (PAGE_WIDTH - PAGE_WIDTH*0.9) / 2  
        pdf.image(img_path, x=x, w=PAGE_WIDTH*0.9)
    else:
        pdf.set_font("DejaVu", "I", 14)
        pdf.set_y(y)
        pdf.cell(
            0,
            20,
            f"Bild '{os.path.basename(img_path)}' fehlt.",
            ln=True,
            align="C",
        )
                  
    # Eckdaten
    pdf.add_page()
#    pages["KEY_DATA"] = int(pdf.get_page_label()) -2 

    page_title(pdf=pdf, csv_data=csv_data, title=text_list["KEY_DATA"].upper())

    list = [
        (text_list["PROPERTY_TYPE"].upper()                   , csv_data["PROPERTY_TYPE"]),
        (text_list["ROOM_FOR_PAYMENT"].upper()                , csv_data["ROOM_FOR_PAYMENT"]),
        (text_list["BEDROOM_FOR_PAYMENT"].upper()             , csv_data["BEDROOM_FOR_PAYMENT"]),
        (text_list["BATHROOM_FOR_PAYMENT"].upper()            , csv_data["BATHROOM_FOR_PAYMENT"]),
        (text_list["GROSS_AREA"].upper()                      , csv_data["GROSS_AREA"]),
        (text_list["RETAIL_SPACE"].upper()                    , csv_data["RETAIL_SPACE"]),
        (text_list["PURCHASE_PRICE_PROPERTY"].upper()         , csv_data["PURCHASE_PRICE"]),
        (text_list["ELEVATOR"].upper()                        , csv_data["ELEVATOR"]),
        (text_list["PARKING_SPACE"].upper()                   , csv_data["PARKING_SPACE"]),
        (text_list["PURCHASE_PRICE_GARAGE"].upper()           , csv_data["PURCHASE_PRICE_GARAGE"]),
        (text_list["FLOOR"].upper()                           , csv_data["FLOOR"]),
        (text_list["TYPE_OF_HEATING"].upper()                 , csv_data["TYPE_OF_HEATING"]),
        (text_list["AIR_CONDITIONING"].upper()                , csv_data["AIR_CONDITIONING"]),
        (text_list["ENERGY_SOURCE"].upper()                   , csv_data["ENERGY_SOURCE"]),
        (text_list["ENERGY_EFFICIENCY_CLASS"].upper()         , csv_data["ENERGY_EFFICIENCY_CLASS"]),
        (text_list["ENERGY_PERFORMANCE_INDEX"].upper()        , csv_data["ENERGY_PERFORMANCE_INDEX"]),
    ]
    list = [(key, value) for key, value in list if value]
    left, right = split_list(list)
    
    pdf.ln(-2)
        
    list_double(pdf, left=left, right=right, line=True)
    pdf.set_y(-150)

    image_files = sorted(
        [
            os.path.join(folder_name, f)
            for f in os.listdir(folder_name)
            if f.startswith("ECK")
            and f.lower().endswith((".jpg", ".jpeg", ".png"))
        ]
    )
    img_path = image_files[0]
    if os.path.exists(img_path):
        # Get image dimensions to maintain aspect ratio
        with Image.open(img_path) as img:
            x = (PAGE_WIDTH - PAGE_WIDTH*0.9) / 2  # Center horizontally
        pdf.image(img_path, x=x, w=PAGE_WIDTH*0.9)
    else:
        pdf.set_font("DejaVu", "I", 14)
        pdf.set_y(y)
        pdf.cell(
            0,
            20,
            f"Bild '{os.path.basename(img_path)}' fehlt.",
            ln=True,
            align="C",
        )
    pdf.add_page()

   # pages["DESCRIPTION"] = int(pdf.get_page_label()) -2 
    page_title(pdf=pdf, csv_data=csv_data, title=text_list["DESCRIPTION"].upper())

    pdf.multi_cell(0, 5, csv_data["DESCRIPTION"], align="L")
    pdf.ln(10)

    pdf.add_page()
    page_title(pdf=pdf, csv_data=csv_data, title=text_list["GOOD_TO_KNOW"].upper())
    
    pdf.set_font("DejaVu", "", 11)
    pdf.multi_cell(0, 5, csv_data["AUSSTATTUNG"], align="L")

    # EINDRÜCKE

    # pages["IMPRESSIONS"] = str(int(int(pdf.get_page_label()) -2 ) + 1)

    image_files = sorted(
        [
            os.path.join(folder_name, f)
            for f in os.listdir(folder_name)
            if f.startswith("EINDRÜCKE_")
            and f.lower().endswith((".jpg", ".jpeg", ".png"))
        ]
    )

    for i in range(0, len(image_files), 2):
        pdf.add_page()
        pdf.set_y(pdf.get_y()-5)
        page_title(pdf=pdf, csv_data=csv_data, title=text_list["IMPRESSIONS"].upper() )
        for j in range(2):
            if i + j < len(image_files):
                img_path = image_files[i + j]
                y = Y_POSITIONS[j]
                if os.path.exists(img_path):
                    # Get image dimensions to maintain aspect ratio
                    with Image.open(img_path) as img:
                        img_width, img_height = img.size
                        aspect_ratio = img_width / img_height
                        img_display_width = IMAGE_HEIGHT * aspect_ratio
                        x = (PAGE_WIDTH - img_display_width) / 2  # Center horizontally

                    pdf.image(img_path, x=x, y=y, h=IMAGE_HEIGHT)
                else:
                    pdf.set_font("DejaVu", "I", 14)
                    pdf.set_y(y)
                    pdf.cell(
                        0,
                        20,
                        f"Bild '{os.path.basename(img_path)}' fehlt.",
                        ln=True,
                        align="C",
                    )

    # LAGEBESCHREIBUNG
    pdf.add_page()
    # pages["LOCATION_DESCRIPTION"] = int(pdf.get_page_label()) -2 
    page_title(pdf=pdf, csv_data=csv_data, title=text_list["LOCATION_DESCRIPTION"].upper())
    pdf.multi_cell(0, 5, csv_data["LOCATION_DESCRIPTION"], align="L")

    pdf.ln(20)
    left = [
        (text_list["HIGHWAY"].upper()   , csv_data["HIGHWAY"]),
        (text_list["CENTER"].upper()    , csv_data["CENTER"]),
        (text_list["AIRPORT"].upper()   , csv_data["AIRPORT"]),
    ]

    list_double(pdf, left=left, right=[], line=True)

    # GRUNDRISSE
    # pages["FLOOR_PLANS"] = int(pdf.get_page_label()) -2 


    # Get image files starting with "EINDRÜCKE_"
    image_files = sorted(
        [
            os.path.join(folder_name, f)
            for f in os.listdir(folder_name)
            if f.startswith("PLAN_")
            and f.lower().endswith((".jpg", ".jpeg", ".png"))
        ]
    )
    TITLE_HEIGHT = 20
    BOTTOM_MARGIN = 15

    # Add one image per page, vertically centered (below title)
    for idx, img_path in enumerate(image_files):
        pdf.add_page()
        pdf.set_y(pdf.get_y() - 5)

        # Draw the title
        page_title(pdf=pdf, csv_data=csv_data, title=text_list["FLOOR_PLANS"].upper())

        # Compute available height and vertical centering
        available_height = pdf.h - TITLE_HEIGHT - BOTTOM_MARGIN
        y_centered = TITLE_HEIGHT + (available_height - IMAGE_HEIGHT) / 2

        if os.path.exists(img_path):
            with Image.open(img_path) as img:
                img_width, img_height = img.size
                aspect_ratio = img_width / img_height
                img_display_width = IMAGE_HEIGHT * aspect_ratio
                x = (PAGE_WIDTH - img_display_width) / 2  # Horizontally center

            pdf.image(img_path, x=x, y=y_centered, h=IMAGE_HEIGHT)
        else:
            pdf.set_font("DejaVu", "I", 14)
            pdf.set_y(y_centered)
            pdf.cell(
                0,
                20,
                f"Bild '{os.path.basename(img_path)}' fehlt.",
                ln=True,
                align="C",
            )



    # FLÄCHENBERECHNUNG IM ÜBERBLICK
    pdf.add_page()
    # pages["AREA_CALCULATION"] = int(pdf.get_page_label()) -2 
    page_title(pdf=pdf, csv_data=csv_data, title=text_list["AREA_CALCULATION"].upper())
    pdf.multi_cell(
        0,
        5,
        text_list["AREA_CALCULATION_TEXT"],
        align="L",
    )

    pdf.set_font("DejaVu", "", 11)
    
    # PROVISION
    pdf.add_page()
    # pages["PROVISION"] = int(pdf.get_page_label()) -2 
    page_title(pdf=pdf, csv_data=csv_data, title=text_list["PROVISION"].upper())
    pdf.multi_cell(
        0,
        5,
        text_list["PROVISION_TEXT"],
        align="L",
    )

    pdf.set_font("DejaVu", "", 11)


    # ANSPRECHPARTNER
    pdf.add_page()
    # pages["CONTACT_PERSON"] = int(pdf.get_page_label()) -2 
    page_title(pdf=pdf, csv_data=csv_data, title=text_list["CONTACT_PERSON"])
    pdf.ln(5)
    pdf.multi_cell(
        0,
        5,
        text_list["CONTACT_PERSON_TEXT_BEFORE"],
        align="L",
    )
    pdf.ln(10)
    y = pdf.get_y()
    pdf.set_xy(100, y)

    pdf.multi_cell(
        90,
        5,
        text_list["CONTACT_PERSON_TEXT"],
        align="L",
    )
    pdf.set_xy(10, y)
    pdf.image("Allgemein/fabian.png", w=80)

    return pdf.output()


# Save PDF
def save_pdf():
    pdf_data = create_pdf()
    file_path = filedialog.asksaveasfilename(
        defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]
    )
    if file_path:
        with open(file_path, "wb") as f:
            f.write(pdf_data)


# Preview PDF
def show_preview():
    global images_tk
    pdf_data = create_pdf()
    images = convert_from_bytes(pdf_data)

    if images:
        for widget in preview_frame.winfo_children():
            widget.destroy()

        images_tk.clear()
        frame_width = max(preview_frame.winfo_width() - 20, 100)
        frame_height = max(preview_frame.winfo_height(), 100)

        for img in images:
            img_copy = img.copy()
            img_copy.thumbnail((frame_width, frame_height))
            img_tk = ImageTk.PhotoImage(img_copy)
            label = tk.Label(preview_frame, image=img_tk)
            label.image = img_tk
            images_tk.append(img_tk)
            label.pack(pady=10)


def on_resize(event):
    # Re-render preview on resize
    show_preview()


def _on_mousewheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

if not valid():
    exit()
    
# --- GUI setup ---
root = tk.Tk()
root.title("Exposé Creator")
root.geometry("1920x1080")

# icon_img = ImageTk.PhotoImage(file="image.png")
# root.iconphoto(True, icon_img)

# Top buttons frame
top_frame = tk.Frame(root)
top_frame.pack(side="top", fill="x")

save_button = tk.Button(top_frame, text="Save Exposé as PDF", command=save_pdf)
save_button.pack(side="left", padx=10, pady=5)

refresh_button = tk.Button(top_frame, text="Refresh Preview", command=show_preview)
refresh_button.pack(side="left", padx=10, pady=5)

def on_language_change(event):
    selected_lang = lang_var.get()
    print(f"Language changed to: {selected_lang}")
    global text_list 
    global lang
    lang = selected_lang
    text_list = load_data_from_csv("Allgemein/DatenAllgemein.csv",selected_lang)

lang_var = tk.StringVar(value="de")
lang_dropdown = ttk.Combobox(top_frame, textvariable=lang_var, values=["de", "it", "en"], state="readonly", width=5)
lang_dropdown.pack(side="left", padx=10, pady=5)
lang_dropdown.bind("<<ComboboxSelected>>", on_language_change)

def get_local_folders(base_path="."):
    excluded = {"Allgemein", "backup", "Barlow-fontiko", "dist", "build", ".git"}
    return [
        f for f in os.listdir(base_path)
        if os.path.isdir(os.path.join(base_path, f)) and f not in excluded
    ]
def on_folder_change(event):
    selected_folder = folder_var.get()
    print(f"Folder selected: {selected_folder}")

    global folder_name
    folder_name =selected_folder


folder_var = tk.StringVar()
folders = get_local_folders(".")  # You can change "." to any path

folder_dropdown = ttk.Combobox(top_frame, textvariable=folder_var, values=folders, state="readonly", width=30)
folder_dropdown.pack(side="left", padx=10, pady=5)
folder_dropdown.bind("<<ComboboxSelected>>", on_folder_change)



# Create canvas
canvas = tk.Canvas(root, width=600, height=700)

scroll_y = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scroll_y.set)

canvas.pack(side="left", fill="both", expand=True)
scroll_y.pack(side="right", fill="y")

preview_frame = tk.Frame(canvas)
preview_window = canvas.create_window((0, 0), window=preview_frame, anchor="nw")
canvas.itemconfig(preview_window, width=canvas.winfo_width())

canvas.configure(scrollregion=canvas.bbox("all"))


# Bind canvas resize
def on_resize(event):
    canvas.itemconfig(preview_window, width=canvas.winfo_width())
    show_preview()


canvas.bind("<Configure>", on_resize)

# Track scrollregion
preview_frame.bind(
    "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

# Bind mousewheel

preview_frame = tk.Frame(canvas)
preview_window = canvas.create_window((0, 0), window=preview_frame, anchor="nw")


# Mouse scroll binding on hover
def _on_mousewheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


root.bind_all("<MouseWheel>", _on_mousewheel)


if platform.system() == "Linux":
    root.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
    root.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))


# Track resizing to adjust previews
canvas.bind("<Configure>", on_resize)

images_tk = []  # Hold references to images

root.after(100, show_preview)
root.mainloop()
