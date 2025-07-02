import tkinter as tk
from tkinter import filedialog
from fpdf import FPDF
from pdf2image import convert_from_bytes
from PIL import ImageTk, Image
import os
import platform
import csv

# ### ToDo
# Bilder automatisch vom folder laden
# fix text abstrahieren
# italienische und englische version vorbereiten
# grfisch switch zwischen sprachen machen


def load_data_from_csv(file_path):
    data = {}
    with open(file_path, newline="", encoding="utf-8") as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            key, value = row
            data[key.strip()] = value.strip()
    return data


pages = {
    "ECKDATEN": 999,
    "BESCHREIBUNG": 999,
    "EINDRÜCKE": 999,
    "LAGEBESCHREIBUNG": 999,
    "GRUNDRISSE": 999,
    "FLÄCHENBEREC": 999,
    "ANSPRECHPARTNER": 999,
}


class ExposePDF(FPDF):
    def __init__(self):
        super().__init__()
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
           self.cell(0, 10, "PREMIUM HOMES", align="L")

           # Page number right-aligned
           self.set_x(-30)
           self.cell(0, 10, f"Seite {self.page_no()-2}", align="R")


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
            pdf.cell(
                60, line_height, left[i][0]
            )  # Left column, first part (no alignment change)
            pdf.set_font("DejaVu", "", 10)
            # Align left[i][1] to the right
            pdf.cell(30, line_height, left[i][1], align="R")

            # draw line under left row
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
            pdf.cell(
                60, line_height, right[i][0]
            )  # Right column, first part (no alignment change)
            pdf.set_font("DejaVu", "", 10)
            # Align right[i][1] to the right
            pdf.cell(30, line_height, right[i][1], align="R")

            # draw line under right row
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
    pdf.ln(20)
    pdf.set_font("DejaVu", "", 11)


def create_pdf():
    pdf = ExposePDF()
    csv_data = load_data_from_csv("Daten.csv")

    # First Page
    pdf.add_page()
    pdf.set_font("DejaVu-TITLE", "", 16)
    pdf.ln(2)
    pdf.image("Allgemein/logo.jpeg", x=(210 - 60) / 2, w=60)
    pdf.ln(5)
    pdf.image(csv_data["TITEL_BILD"], x=(210 - 180) / 2, w=180)
    pdf.ln(2)
    pdf.multi_cell(
        0,
        8,
        csv_data["TITEL"],
        align="C",
    )
    pdf.ln(5)
    left = [
        ("ART", csv_data["OBJEKTTYP"]),
        ("ANZAHL ZIMMER", csv_data["ANZAHLZIMMER"]),
    ]

    right = [
        ("HANDELSFLÄCHE", csv_data["HANDELSFLÄCHE"]),
        ("KAUFPREIS", csv_data["KAUFPREISIMMOBILIE"]),
    ]

    list_double(pdf, left=left, right=right, line=True)
    pdf.ln(40)
    pdf.set_font("DejaVu", "B", 11)
    pdf.set_y(-40)
    pdf.cell(0, 5, "FABIAN PERNTHALER", ln=True, align="L")

    pdf.set_font("DejaVu", "", 11)
    pdf.cell(0, 5, "f.pernthaler@premium-homes.it", ln=False, align="L")
    pdf.cell(0, 5, "www.premium-homes.it", ln=True, align="R")
    pdf.cell(0, 5, "+39 347 4734794 | DE, IT, EN", ln=False, align="L")
    pdf.cell(0, 5, "@premiumhomes.lakegarda", ln=True, align="R")

    # Table of Contents
    pdf.add_page()
    page_title(pdf=pdf, csv_data=csv_data, title="ÜBERSICHT")

    left = [
        ("ECKDATEN", "Seite" + pages["ECKDATEN"].__str__()),
        ("BESCHREIBUNG", "Seite" + pages["BESCHREIBUNG"].__str__()),
        ("EINDRÜCKE", "Seite" + pages["EINDRÜCKE"].__str__()),
        ("LAGEBESCHREIBUNG", "Seite" + pages["LAGEBESCHREIBUNG"].__str__()),
        ("GRUNDRISSE", "Seite" + pages["GRUNDRISSE"].__str__()),
        ("FLÄCHENBERECHNUNG IM ÜBERBLICK", "Seite" + pages["FLÄCHENBEREC"].__str__()),
        ("ANSPRECHPARTNER", "Seite" + pages["ANSPRECHPARTNER"].__str__()),
    ]

    list_single(pdf, left=left, line=True)

    pdf.set_y(-170)
    folder_path = csv_data["EINDRUCK_ORDNER"]

    # Constants
    PAGE_WIDTH = 210  # A4 width in mm

    # Get image files starting with "EINDRÜCKE_"
    image_files = sorted(
        [
            os.path.join(folder_path, f)
            for f in os.listdir(folder_path)
            if f.startswith("ÜBERSICHT")
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
                  
    # Eckdaten
    pdf.add_page()
    pages["ECKDATEN"] = int(pdf.get_page_label()) -2 

    page_title(pdf=pdf, csv_data=csv_data, title="Eckdaten")

    left = [
        ("OBJEKTTYP", csv_data["OBJEKTTYP"]),
        ("ANZAHL ZIMMER", csv_data["ANZAHLZIMMER"]),
        ("ANZAHL SCHLAFZIMMER", csv_data["ANZAHLSCHLAFZIMMER"]),
        ("ANZAHL BADEZIMMER", csv_data["ANZAHLBADEZIMMER"]),
        ("BRUTTOFLÄCHE", csv_data["BRUTTOFLÄCHE"]),
        ("HANDELSFLÄCHE", csv_data["HANDELSFLÄCHE"]),
        ("KAUFPREIS IMMOBILIE", csv_data["KAUFPREISIMMOBILIE"]),
        ("FAHRSTUHL", csv_data["FAHRSTUHL"]),
    ]

    right = [
        ("PKW STELLFLÄCHEN", csv_data["PKWSTELLFLÄCHEN"]),
        ("KAUFPREIS GARAGE", csv_data["KAUFPREISGARAGE"]),
        ("ETAGE", csv_data["ETAGE"]),
        ("HEIZUNGSART", csv_data["HEIZUNGSART"]),
        ("KLIMATISIERUNG", csv_data["KLIMATISIERUNG"]),
        ("ENERGIETRÄGER", csv_data["ENERGIETRÄGER"]),
        ("ENERGIEEFFIZIENZKLASSE", csv_data["ENERGIEEFFIZIENZKLASSE"]),
        ("ENERGIELEISTUNGSINDEX", csv_data["ENERGIELEISTUNGSINDEX"]),
    ]

    list_double(pdf, left=left, right=right, line=True)
    pdf.set_y(-170)
    folder_path = csv_data["EINDRUCK_ORDNER"]

    # Constants
    PAGE_WIDTH = 210  # A4 width in mm

    # Get image files starting with "EINDRÜCKE_"
    image_files = sorted(
        [
            os.path.join(folder_path, f)
            for f in os.listdir(folder_path)
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

    pages["BESCHREIBUNG"] = int(pdf.get_page_label()) -2 
    page_title(pdf=pdf, csv_data=csv_data, title="BESCHREIBUNG")

    pdf.multi_cell(0, 5, csv_data["BESCHREIBUNG"], align="L", markdown=True)
    pdf.ln(10)

    pdf.add_page()
    page_title(pdf=pdf, csv_data=csv_data, title="GUT ZU WISSEN")
    
    pdf.set_font("DejaVu", "", 11)
    pdf.multi_cell(0, 5, csv_data["AUSSTATTUNG"], align="L")

    # EINDRÜCKE

    pages["EINDRÜCKE"] = str(int(int(pdf.get_page_label()) -2 ) + 1)
    folder_path = csv_data["EINDRUCK_ORDNER"]

    # Constants
    PAGE_WIDTH = 210  # A4 width in mm
    PAGE_HEIGHT = 297  # A4 height in mm
    IMAGE_HEIGHT = PAGE_HEIGHT * 0.43  # 40% = ~119mm
    Y_POSITIONS = [20, 150]  # Y positions for the 2 images

    # Get image files starting with "EINDRÜCKE_"
    image_files = sorted(
        [
            os.path.join(folder_path, f)
            for f in os.listdir(folder_path)
            if f.startswith("EINDRÜCKE_")
            and f.lower().endswith((".jpg", ".jpeg", ".png"))
        ]
    )

    # Add images to PDF, two per page
    for i in range(0, len(image_files), 2):
        pdf.add_page()
        pdf.set_y(pdf.get_y()-5)
        page_title(pdf=pdf, csv_data=csv_data, title="EINDRÜCKE")
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
    pages["LAGEBESCHREIBUNG"] = int(pdf.get_page_label()) -2 
    page_title(pdf=pdf, csv_data=csv_data, title="LAGEBESCHREIBUNG")
    pdf.multi_cell(0, 5, csv_data["LAGEBESCHREIBUNG"], align="L")

    pdf.ln(20)
    left = [
        ("DIST. AUTOBAHN (KM)", csv_data["DIST.AUTOBAHN(KM)"]),
        ("DIST. ZENTRUM (KM)", csv_data["DIST.ZENTRUM(KM)"]),
        ("DIST. FLUGHAFEN (KM)", csv_data["DIST.FLUGHAFEN(KM)"]),
    ]

    list_double(pdf, left=left, right=[], line=True)

    # GRUNDRISSE
    pages["GRUNDRISSE"] = int(pdf.get_page_label()) -2 

    folder_path = csv_data["EINDRUCK_ORDNER"]

    # Constants
    PAGE_WIDTH = 210  # A4 width in mm
    PAGE_HEIGHT = 297  # A4 height in mm
    IMAGE_HEIGHT = PAGE_HEIGHT * 0.43  # 40% = ~119mm
    Y_POSITIONS = [20, 150]  # Y positions for the 2 images

    # Get image files starting with "EINDRÜCKE_"
    image_files = sorted(
        [
            os.path.join(folder_path, f)
            for f in os.listdir(folder_path)
            if f.startswith("PLAN_")
            and f.lower().endswith((".jpg", ".jpeg", ".png"))
        ]
    )

    # Add images to PDF, two per page
    for i in range(0, len(image_files), 2):
        pdf.add_page()
        pdf.set_y(pdf.get_y()-5)
        page_title(pdf=pdf, csv_data=csv_data, title="GRUNDRISSE")
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





    # FLÄCHENBERECHNUNG IM ÜBERBLICK
    pdf.add_page()
    pages["FLÄCHENBEREC"] = int(pdf.get_page_label()) -2 
    page_title(pdf=pdf, csv_data=csv_data, title="FLÄCHENBERECHNUNG IM ÜBERBLICK")

    pdf.multi_cell(
        0,
        5,
        """Die Handelsfläche, auch Verkaufsfläche genannt (im italienischen "Superficie commerciale") ist eine standardisierte Berechnungsmethode für den Verkauf von Immobilien in Italien, die über die reine Wohnfläche hinausgeht und in großteils der Immobilienportalen, Werbungen und Exposés Anwendung findet.""",
    )
    pdf.set_font("DejaVu", "B", 11)
    pdf.ln(10)
    pdf.cell(0, 10, "INNENRÄUME", ln=True)
    pdf.set_font("DejaVu", "", 11)
    list_items = [
        "Vollständige Berechnung der Wohnfläche inkl. Innen- und Außenwände",
        "Gemeinsame Trennwände werden zur Hälfte eingerechnet",
    ]

    # Adding a bullet point list
    for item in list_items:
        pdf.cell(0, 5, f"• {item}", ln=True)

    pdf.cell(
        0,
        10,
        "Nebenflächen werden anteilig mit reduzierten Prozentsätzen berechnet:",
        ln=True,
    )
    list_items = [
        "Keller: ca. 20–35 %",
        "Balkone/Terrassen: ca. 25–50%",
        "Gärten: ca. 10–15 %",
        "Garagen: ca. 50 %, nicht überdachte Stellplätze ca. 20%",
        "Gemeinsame Trennwände werden zur Hälfte eingerechnet",
    ]
    # Adding a bullet point list
    for item in list_items:
        pdf.cell(0, 5, f"• {item}", ln=True)
    # pdf.ln(20)

    pdf.ln(10)
    pdf.set_font("DejaVu", "B", 11)
    pdf.cell(0, 10, "BRUTTOFLÄCHE (SUPERFICIE LORDA)", ln=True)
    pdf.set_font("DejaVu", "", 11)
    pdf.multi_cell(
        0,
        5,
        """Die Bruttofläche ist die Gesamtfläche aller Innenräume einer Immobilie, sowohl der bewohnbaren als auch der nicht bewohnbaren Flächen. Die nicht bewohnbaren Flächen, wie Kellergeschosse werden meistens nur zu einem Anteil, wie etwa einem Drittel einberechnet. Die Bruttofläche umfasst die Innenwände und Außenwände, schließt jedoch sekundäre Außenflächen wie Balkone, Terrassen und Gärten aus.
                   
Detaillierte Informationen zur Berechnung der Verkaufsfläche finden Sie in einem unserer Blogartikel auf --www.premium-homes.it--. Zusätzlich bietet unser Glossar umfassende Erklärungen zur Verkaufs- und Bruttofläche in Italien.""",
        markdown=True,
    )

    pdf.set_font("DejaVu", "", 11)

    # ANSPRECHPARTNER
    pdf.add_page()
    pages["ANSPRECHPARTNER"] = int(pdf.get_page_label()) -2 
    page_title(pdf=pdf, csv_data=csv_data, title="ANSPRECHPARTNER")
    pdf.ln(10)
    pdf.multi_cell(
        0,
        5,
        """Weitere Informationen erhalten Sie über Ihren Ansprechpartner und eingetragenen Immobilienmakler Fabian Pernthaler.""",
        markdown=True,
        align="L",
    )
    pdf.ln(10)
    y = pdf.get_y()
    pdf.set_xy(100, y)

    pdf.multi_cell(
        90,
        5,
        """**FABIAN PERNTHALER**
--f.pernthaler@premium-homes.it--
+39 347 4734794
Sprachen | DE IT EN
        
Mit über sechs Jahren Erfahrung als Immobilienmakler und einem Bachelorabschluss in Facility Management & Immobilienwirtschaft bringe ich fundiertes Fachwissen und Leidenschaft für alles mit, was mit Immobilien und Menschen zu tun hat. Darüber hinaus habe ich durch meine frühere Arbeit in einem Architekturbüro wertvolle Einblicke in planerische und bauliche Aspekte gewonnen – entscheidende Details, die oft den Unterschied ausmachen.
        
Ich freue mich darauf, über Ihr Vorhaben zu sprechen.""",
        markdown=True,
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
