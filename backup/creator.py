import tkinter as tk
from tkinter import filedialog
from fpdf import FPDF
from pdf2image import convert_from_bytes
from PIL import ImageTk, Image
import os
import platform


# Your image filenames (replace with real images later!)
IMAGE_FILES = ["img1.png", "img2.png", "img3.png", "img4.png", "img5.png"]


class ExposePDF(FPDF):
    def __init__(self):
        super().__init__()
        self.add_font(
            "DejaVu", "", "./dejavu-sans/DejaVuSans.ttf", uni=True
        )  # <-- Unicode font
        self.add_font(
            "DejaVu", "I", "./dejavu-sans/DejaVuSansCondensed-Oblique.ttf", uni=True
        )  # Italic font
        self.add_font(
            "DejaVu", "B", "./dejavu-sans/DejaVuSans-Bold.ttf", uni=True
        )  # Italic font

        self.set_font("DejaVu", "", 14)

    def header(self):
        pass  # No default header

    def footer(self):
        self.set_y(-15)  # Position 15mm from bottom
        self.set_font("DejaVu", "I", 8)

        # Left-aligned "PREMIUM HOMES"
        self.cell(0, 10, "PREMIUM HOMES", align="L")

        # Page number right-aligned
        self.set_x(-30)
        self.cell(0, 10, f"Seite {self.page_no()}", align="R")


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
            pdf.set_font("DejaVu", "B", 10)
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
                    y + line_height,
                    start_x_left + col_width,
                    y + line_height,
                )

        if i < len(right):
            pdf.set_xy(start_x_right, y)
            pdf.set_font("DejaVu", "B", 10)
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
                    y + line_height,
                    start_x_right + col_width,
                    y + line_height,
                )


def create_pdf():
    pdf = ExposePDF()

    # First Page
    pdf.add_page()
    pdf.set_font("DejaVu", "B", 16)
    pdf.image("logo.png", x=70, w=80)
    pdf.image("titel_bild.png", x=20, w=160)
    pdf.multi_cell(
        0,
        10,
        "MODERNE ATTIKA-WOHNUNG MIT GROSSER\nDACHTERRASSE IN UNMITTELBARER SEENÄHE",
        align="C",
    )

    left = [
        ("ART", "Wohnung"),
        ("ANZAHL ZIMMER", "3"),
    ]

    right = [
        ("HANDELSFLÄCHE", "150 m²"),
        ("KAUFPREIS", "575.000 €"),
    ]

    list_double(pdf, left=left, right=right, line=True)
    pdf.ln(20)
    pdf.set_font("DejaVu", "B", 11)

    pdf.cell(0, 5, "FABIAN PERNTHALER", ln=True, align="L")

    pdf.set_font("DejaVu", "", 11)
    pdf.cell(0, 5, "f.pernthaler@premium-homes.it", ln=False, align="L")
    pdf.cell(0, 5, "www.premium-homes.it", ln=True, align="R")
    pdf.cell(0, 5, "+39 347 4734794 | DE, IT, EN", ln=False, align="L")
    pdf.cell(0, 5, "@premiumhomes.lakegarda", ln=True, align="R")

    # Table of Contents
    pdf.add_page()
    pdf.set_font("DejaVu", "B", 18)
    pdf.cell(0, 10, "Eckdaten", ln=True)
    pdf.set_font("DejaVu", "", 16)
    pdf.cell(0, 10, "PH160 - Desenzano del Garda", ln=True)
    pdf.ln(10)
    pdf.set_font("DejaVu", "", 12)
    pdf.cell(
        0,
        10,
        "1. Exposé Übersicht ....................................... Seite 1",
        ln=True,
    )
    pdf.cell(
        0,
        10,
        "2. Inhaltsverzeichnis .................................. Seite 2",
        ln=True,
    )
    pdf.cell(
        0,
        10,
        "3. Beschreibung .......................................... Seite 3",
        ln=True,
    )
    pdf.cell(
        0,
        10,
        "4. Eindrücke ................................................. Seite 4-8",
        ln=True,
    )

    # Table of Contents
    pdf.add_page()
    pdf.set_font("DejaVu", "B", 18)
    pdf.cell(0, 10, "Eckdaten", ln=True, align="C")
    pdf.set_font("DejaVu", "", 16)
    pdf.cell(0, 10, "PH160 - Desenzano del Garda", ln=True, align="C")
    pdf.ln(10)

    left = [
        ("OBJEKTTYP", "Wohnung"),
        ("ANZAHL ZIMMER", "3"),
        ("ANZAHL SCHLAFZIMMER", "2"),
        ("ANZAHL BADEZIMMER", "2"),
        ("BRUTTOFLÄCHE", "90 m²"),
        ("HANDELSFLÄCHE", "150 m²"),
        ("KAUFPREIS IMMOBILIE", "575.000 EUR"),
        ("FAHRSTUHL", "ja"),
    ]

    right = [
        ("PKW STELLFLÄCHEN", "Doppelgarage"),
        ("KAUFPREIS GARAGE", "Inkludiert"),
        ("ETAGE", "2. Etage"),
        ("HEIZUNGSART", "Fußbodenheizung"),
        ("KLIMATISIERUNG", "moderne Splitgeräte"),
        ("ENERGIETRÄGER", "Wärmepumpe"),
        ("ENERGIEEFFIZIENZKLASSE", "Legge 90/2013, A"),
        ("ENERGIELEISTUNGSINDEX", "72,82 kWh/m²a"),
    ]

    list_double(pdf, left=left, right=right, line=True)

    pdf.add_page()
    pdf.set_font("DejaVu", "B", 18)
    pdf.cell(0, 10, "BESCHREIBUNG", ln=True, align="C")
    pdf.set_font("DejaVu", "", 16)
    pdf.cell(0, 10, "PH160 - Desenzano del Garda", ln=True, align="C")
    pdf.ln(10)

    pdf.set_font("DejaVu", "", 11)
    pdf.multi_cell(
        0,
        5,
        """Erleben Sie hochwertiges Wohnen in dieser modernen Attika-Wohnung in einer ansprechenden und kürzlich erbauten Residenz in unmittelbarer Seenähe von Rivoltella.

Diese neuwertige Wohnung bietet höchsten Wohnkomfort und besticht durch Ihre großzügigen Außenflächen, den riesigen Pool und der hervorragenden Lage.

Die Wohnung verfügt über ein großzügiges Wohnzimmer mit offener Küche und Zugang zur Loggia, ideal für gesellige Abende und entspannte Stunden. Das Hauptschlafzimmer ist mit einem eigenen Bad en Suite ausgestattet, während ein zweites Schlafzimmer und ein weiteres Bad zusätzlichen Komfort bieten. Ein weiterer Balkon mit anliegenden Abstellraum und ein Atrium, von dem aus eine Treppe zur großzügigen Dachterrasse führt, vervollständigen das Raumangebot. Auf der Dachterrasse genießen Sie einen teilweisen Seeblick und die umliegende Naturlandschaft, die zum Entspannen und Verweilen einlädt. Zusätzlich besteht die Möglichkeit, einen Whirlpool auf der Dachterrasse anzubringen.

Die Wohnung wurde in der Energieklasse A3 realisiert und ist mit modernsten Technologien ausgestattet. Dazu gehören Fußbodenheizung, Klimatisierung in allen Wohn- und Schlafbereichen, eine Photovoltaikanlage und automatische Wohnraumlüftung. Elektrische Rollläden und ein Aufzug sorgen für zusätzlichen Komfort und Sicherheit.

Zur Wohnung gehört eine große Garage mit Platz für zwei Autos und zusätzlicher Staufläche, welche im Kaufpreis inkludiert ist. Dies bietet nicht nur Sicherheit für Ihre Fahrzeuge, sondern auch zusätzlichen Stauraum für Ihre persönlichen Gegenstände. 

Alle wichtigen Dienstleistungen sind in unmittelbarer Nähe vorhanden. Ein Strand ist fußläufig in nur 10 Minuten erreichbar, und der See ist nur 200 Meter Luftlinie entfernt. 

Die 10 % Mehrwertsteuer als Zweitwohnsitz auf den Kaufpreis entfallen bei dieser Wohnung, da es sich um einen Privatverkauf handelt und nicht um einen Erwerb direkt vom Bauträger. 

Diese Attika-Wohnung ist die perfekte Wahl für alle, die modernen Komfort, großzügige Außenflächen und eine erstklassige Lage zu schätzen wissen. Ob als Hauptwohnsitz oder als Ferienwohnung - diese Immobilie bietet Ihnen alles, was Sie sich wünschen.

Vereinbaren Sie noch heute einen Besichtigungstermin und überzeugen Sie sich selbst von dieser außergewöhnlichen Wohnung!""",
    )
    pdf.ln(10)
    pdf.set_font("DejaVu", "B", 14)
    pdf.cell(0, 10, "AUSSTATTUNG", ln=True, align="L")
    pdf.ln(10)
    pdf.set_font("DejaVu", "", 11)
    pdf.multi_cell(
        0,
        5,
        """- Die Einbauküche ist im Kaufpreis nicht inkludiert und kann für 20.000EUR abgelöst werden. Die Küche wurde hochwertig in Granitplatte, Herd mit Abzug, Kühl/Gefrierkombi, Backofen mit integriertem Dampfgarer, realisiert.
- Fußbodenheizung, Klimatisierung in allen Wohn- und Schlafbereichen, Photovoltaikanlage, automatische Wohnraumlüftung und mit Vorrichtungen für eine Alarmanlage und für Smart
- Home, sowie elektrische Rolläden.
- Aufzug.
- Videosprechanlage.
- 25m langer Pool mit Terrasse und gepflegter Gartenanlage.
- Glasfaser Anschluss.""",
    )

    # Picture Pages
    for img_file in IMAGE_FILES:
        if os.path.exists(img_file):
            pdf.add_page()
            pdf.image(img_file, x=10, y=20, w=190)  # Adjust width/height
        else:
            # Fallback if image not found
            pdf.add_page()
            pdf.set_font("DejaVu", "I", 14)
            pdf.cell(0, 20, f"Bild '{img_file}' fehlt.", ln=True, align="C")

    # LAGEBESCHREIBUNG
    pdf.add_page()
    pdf.set_font("DejaVu", "B", 18)
    pdf.cell(0, 10, "LAGEBESCHREIBUNG", ln=True, align="C")
    pdf.set_font("DejaVu", "", 16)
    pdf.cell(0, 10, "PH160 - Desenzano del Garda", ln=True, align="C")
    pdf.ln(10)

    pdf.set_font("DejaVu", "", 11)
    pdf.multi_cell(
        0,
        5,
        """Diese traumhafte Attika Wohnung befindet sich in Rivoltella, einem charmanten Ortsteil von Desenzano del Garda, nur wenige Schritte vom Ufer des Gardasees entfernt. Die Lage vereint auf ideale Weise Ruhe, Naturnähe und erstklassige Anbindung an zwei der beliebtesten Orte am südlichen Gardasee: Desenzano und Sirmione. 
                   
Rivoltella besticht durch seine entspannte Atmosphäre, die Nähe zum Wasser und eine gute Infrastruktur mit Restaurants, Cafés, kleinen Läden und Wochenmarkt. Der hübsche kleine Hafen, die Seepromenade sowie die Nähe zu gepflegten Stränden machen den Ort besonders attraktiv für Feriengäste und Liebhaber des "dolce vita". 
                   
Nur wenige Minuten westlich liegt Desenzano del Garda - das lebendige Herz des südlichen Gardasees. Hier erwartet dich ein ganzjährig belebter Ort mit eleganten Boutiquen, einer einladenden Altstadt, Yachthafen, zahlreichen Bars, Restaurants und direktem Bahnanschluss nach Mailand und Verona. Desenzano ist eine perfekte Mischung aus italienischer Lebensfreude, Kultur und moderner Infrastruktur. 
                   
Östlich von Rivoltella befindet sich die berühmte Halbinsel Sirmione, die zu den bekanntesten Reisezielen am Gardasee gehört. Die Altstadt von Sirmione mit ihrer historischen Scaligerburg, den römischen Grotten des Catull und den charmanten Gassen zieht Besucher aus aller Welt an. Neben kulturellen Highlights bietet Sirmione auch erstklassige Thermen, feine Gastronomie und eine zauberhafte Atmosphäre direkt am Wasser. 
                   
Dank der zentralen Lage ist Rivoltella der ideale Ausgangspunkt, um sowohl das lebhafte Desenzano als auch das elegante Sirmione schnell und bequem zu erreichen - und gleichzeitig die Ruhe einer gepflegten Wohnlage am See zu genießen. Diese Immobilie ist daher nicht nur eine erstklassige Wahl für die Eigennutzung, sondern auch eine hervorragende Investition im Premiumsegment des Gardasees.""",
    )

    pdf.ln(20)
    left = [
        ("DIST. AUTOBAHN (KM)", "4 km"),
        ("DIST. ZENTRUM (KM)", "1,2 km"),
        ("DIST. FLUGHAFEN (KM)", "30 km"),
    ]

    list_double(pdf, left=left, right=[], line=True)

    # GRUNDRISSE

    # FLÄCHENBERECHNUNG IM ÜBERBLICK

    pdf.add_page()
    pdf.set_font("DejaVu", "B", 18)
    pdf.cell(0, 10, "FLÄCHENBERECHNUNG IM ÜBERBLICK", ln=True, align="C")
    pdf.set_font("DejaVu", "", 16)
    pdf.cell(0, 10, "PH160 - Desenzano del Garda", ln=True, align="C")
    pdf.ln(10)

    pdf.set_font("DejaVu", "", 11)
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
    pdf.set_font("DejaVu", "B", 18)
    pdf.cell(0, 10, "ANSPRECHPARTNER", ln=True, align="C")
    pdf.set_font("DejaVu", "", 16)
    pdf.cell(0, 10, "PH160 - Desenzano del Garda", ln=True, align="C")

    pdf.set_font("DejaVu", "", 11)
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
    pdf.image("fabian.png", w=80)
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

icon_img = ImageTk.PhotoImage(file="image.png")
root.iconphoto(True, icon_img)

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

# Bind canvas resize
def on_resize(event):
    canvas.itemconfig(preview_window, width=canvas.winfo_width())
    show_preview()

canvas.bind("<Configure>", on_resize)

# Track scrollregion
preview_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

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