"""Generates the fictitious sample documents in samples/documents.

Everything below is invented: the regulator, the case, the people, the
addresses and the email domains (.example). The documents cover the inputs the
bundling pipeline handles: text PDFs (A4 and US Letter), a scanned PDF without
a text layer, a JPEG photo of a handwritten note, Word documents, a Word cover
sheet template with placeholders, and .eml emails with attachments.

Usage (from the repository root, in a virtual environment outside the repo):
    python3 -m venv /tmp/bundling-samples-venv
    /tmp/bundling-samples-venv/bin/pip install -r scripts/requirements.txt
    /tmp/bundling-samples-venv/bin/python scripts/make-samples.py
"""

import io
import random
from datetime import datetime, timezone
from email.message import EmailMessage
from email.policy import SMTP
from email.utils import format_datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm, Pt
from PIL import Image, ImageDraw, ImageFilter, ImageFont
from reportlab.lib.pagesizes import A4, LETTER
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

OUT = Path(__file__).resolve().parent.parent / "samples" / "documents"

REGULATOR = "Example Professional Standards Council"
REGULATOR_SHORT = "EPSC"
CASE_REF = "EPSC-2026-0142"
REGISTRANT = "Jordan Ellis"
REGISTRATION = "PH-204518"
PHARMACY = "Riverside Community Pharmacy, 12 Mill Road, Northbridge NB4 2QT"
FIXED_DATE = datetime(2026, 3, 2, 9, 0, tzinfo=timezone.utc)


# ---------------------------------------------------------------- fonts
def font(size, candidates):
    """First TrueType font found on this machine, else Pillow's built-in font."""
    for name in candidates:
        for folder in ("/System/Library/Fonts/Supplemental", "/System/Library/Fonts", "/Library/Fonts",
                       "/usr/share/fonts/truetype/dejavu", "/usr/share/fonts/truetype/liberation",
                       "C:/Windows/Fonts"):
            path = Path(folder) / name
            if path.exists():
                return ImageFont.truetype(str(path), size)
    return ImageFont.load_default(size=size)


HAND = ["Bradley Hand Bold.ttf", "Noteworthy.ttc", "MarkerFelt.ttc", "DejaVuSans.ttf", "segoepr.ttf", "arial.ttf"]
TYPEWRITER = ["Courier New.ttf", "Courier.ttc", "DejaVuSansMono.ttf", "LiberationMono-Regular.ttf", "cour.ttf"]


# ---------------------------------------------------------------- PDF helpers
class Pdf:
    """Small wrapper around a reportlab canvas that writes lines top-down."""

    def __init__(self, path_or_buffer, pagesize=A4, title=""):
        self.c = canvas.Canvas(path_or_buffer, pagesize=pagesize, invariant=1)
        self.c.setTitle(title)
        self.c.setAuthor(REGULATOR)
        self.c.setCreator("make-samples.py")
        self.w, self.h = pagesize
        self.y = self.h - 25 * mm

    def letterhead(self, subtitle):
        c = self.c
        c.setFillColorRGB(0.13, 0.27, 0.45)
        c.rect(0, self.h - 18 * mm, self.w, 18 * mm, stroke=0, fill=1)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica-Bold", 13)
        c.drawString(20 * mm, self.h - 11 * mm, REGULATOR)
        c.setFont("Helvetica", 8.5)
        c.drawRightString(self.w - 20 * mm, self.h - 11 * mm, subtitle)
        c.setFillColorRGB(0, 0, 0)
        self.y = self.h - 32 * mm

    def text(self, line, size=10.5, bold=False, gap=None, indent=0):
        if self.y < 25 * mm:
            self.new_page()
        self.c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        self.c.drawString(20 * mm + indent, self.y, line)
        self.y -= gap if gap is not None else size * 1.5

    def para(self, text, size=10.5, width=95):
        """Wraps a paragraph at about `width` characters."""
        words, line = text.split(), ""
        for word in words:
            if len(line) + len(word) + 1 > width:
                self.text(line, size)
                line = word
            else:
                line = f"{line} {word}".strip()
        if line:
            self.text(line, size)
        self.y -= size * 0.6

    def field(self, label, value):
        self.c.setFont("Helvetica-Bold", 9.5)
        self.c.drawString(20 * mm, self.y, label)
        self.c.setFont("Helvetica", 10.5)
        self.c.drawString(70 * mm, self.y, value)
        self.c.setStrokeColorRGB(0.7, 0.7, 0.7)
        self.c.line(68 * mm, self.y - 2.5, self.w - 20 * mm, self.y - 2.5)
        self.y -= 19

    def new_page(self):
        self.c.showPage()
        self.y = self.h - 25 * mm

    def save(self):
        self.c.save()


# ---------------------------------------------------------------- PDF documents
def referral_form():
    pdf = Pdf(str(OUT / "Referral form.pdf"), title="Referral form")
    pdf.letterhead("Fitness to practise referral form")
    pdf.text("Referral form", 16, bold=True, gap=26)
    pdf.text("Section 1 - About the registrant", 11.5, bold=True, gap=20)
    pdf.field("Name", REGISTRANT)
    pdf.field("Registration number", REGISTRATION)
    pdf.field("Role", "Pharmacist (responsible pharmacist on duty)")
    pdf.field("Place of work", "Riverside Community Pharmacy, Northbridge")
    pdf.y -= 8
    pdf.text("Section 2 - About the referrer", 11.5, bold=True, gap=20)
    pdf.field("Name", "Chris Walker")
    pdf.field("Role", "Pharmacy manager")
    pdf.field("Email", "c.walker@riverside-pharmacy.example")
    pdf.field("Date of referral", "2 February 2026")
    pdf.y -= 8
    pdf.text("Section 3 - Summary of concerns", 11.5, bold=True, gap=20)
    pdf.para("On 27 January 2026 a patient received 50 mg tablets instead of the 5 mg tablets that were "
             "prescribed. The error was noticed by the patient's carer the next day and no harm was reported. "
             "A review of the controlled drugs register for January found three entries that were not made on "
             "the day of supply. The referrer is concerned about dispensing checks and record keeping.")
    pdf.text("Section 4 - Declaration", 11.5, bold=True, gap=20)
    pdf.para("I confirm that the information in this form is true to the best of my knowledge.")
    pdf.field("Signed", "C. Walker")
    pdf.save()


def incident_report():
    # US Letter on purpose: the pipeline normalises every page to A4.
    pdf = Pdf(str(OUT / "CW02 incident report.pdf"), pagesize=LETTER, title="Incident report")
    pdf.letterhead("Exhibit CW/02")
    pdf.text("Riverside Community Pharmacy - Incident report", 15, bold=True, gap=24)
    pdf.field("Incident number", "RCP-INC-2026-007")
    pdf.field("Date and time", "27 January 2026, 16:40")
    pdf.field("Reported by", "Chris Walker, pharmacy manager")
    pdf.field("Staff involved", f"{REGISTRANT} (responsible pharmacist), one dispensing assistant")
    pdf.y -= 6
    pdf.text("What happened", 11.5, bold=True, gap=18)
    pdf.para("A repeat prescription for 5 mg tablets was dispensed as 50 mg tablets. Both strengths are stored "
             "next to each other on the same shelf. The bag was checked and handed out at 16:40. On 28 January "
             "the patient's carer called the pharmacy after reading the box label. The patient had not taken "
             "any of the tablets.")
    pdf.text("Immediate actions", 11.5, bold=True, gap=18)
    for line in ["- The correct medicine was delivered to the patient's home on 28 January at 11:15.",
                 "- The two strengths were moved to separate shelves with shelf-edge warnings.",
                 "- The incident was recorded on the pharmacy's patient safety system.",
                 "- The patient's GP practice was informed by phone and by email."]:
        pdf.text(line, 10.5, indent=4)
    pdf.y -= 6
    pdf.text("Contributing factors", 11.5, bold=True, gap=18)
    pdf.para("The pharmacy was short of one member of staff that afternoon. The second check was done by the "
             "same pharmacist who labelled the item. The dispensing system showed a stock warning that was "
             "dismissed.")
    pdf.new_page()
    pdf.text("Learning and follow-up", 11.5, bold=True, gap=18)
    pdf.para("A two-person check is now required for all look-alike, sound-alike medicines. The team will "
             "review the stock warning process at the next monthly meeting. This report was shared with the "
             "superintendent pharmacist on 30 January 2026.")
    pdf.field("Completed by", "Chris Walker")
    pdf.field("Date", "30 January 2026")
    pdf.save()


def investigation_report():
    pdf = Pdf(str(OUT / "PS01 investigation report.pdf"), title="Investigation report")
    pdf.letterhead("Exhibit PS/01")
    pdf.text("Investigation report", 16, bold=True, gap=22)
    pdf.text(f"Case reference {CASE_REF}  |  Registrant: {REGISTRANT} ({REGISTRATION})", 9.5, gap=24)
    sections = [
        ("1. Scope", "This report covers the dispensing error of 27 January 2026 and the record keeping of the "
                     "controlled drugs register at Riverside Community Pharmacy between 1 and 31 January 2026."),
        ("2. Evidence reviewed", "The referral form, the pharmacy incident report (exhibit CW/02), a handwritten "
                                 "note by the pharmacy manager (exhibit CW/01), the controlled drugs register, "
                                 "the staff rota for January and the investigator's site visit notes "
                                 "(exhibit PS/02)."),
        ("3. Dispensing error", "The labelling and the final check were done by the same person. The rota shows "
                                "that only two staff were on duty from 14:00 to 18:00 against a planned three. "
                                "The registrant accepts the error and reported it to the manager the next "
                                "morning."),
        ("4. Controlled drugs register", "Three entries for January were written one to two days after the supply. "
                                         "The quantities and balances are correct. No stock is missing."),
        ("5. Remediation", "The registrant completed a refresher course on safe dispensing on 20 February 2026 "
                           "and the pharmacy introduced a two-person check for look-alike medicines."),
        ("6. Conclusion", "The facts are not disputed. The panel is asked to decide whether the registrant's "
                          "fitness to practise is currently impaired, taking into account the remediation "
                          "described above."),
    ]
    for heading, body in sections:
        pdf.text(heading, 11.5, bold=True, gap=18)
        pdf.para(body)
        if heading.startswith("3."):
            pdf.new_page()
    pdf.field("Investigator", "Priya Shah")
    pdf.field("Date", "24 February 2026")
    pdf.save()


def notice_of_hearing():
    buffer = io.BytesIO()
    pdf = Pdf(buffer, title="Notice of hearing")
    pdf.letterhead("Hearings team")
    pdf.text("Notice of hearing", 16, bold=True, gap=26)
    pdf.text(f"To: {REGISTRANT}", 10.5)
    pdf.text(f"Case reference: {CASE_REF}", 10.5, gap=24)
    pdf.para(f"The {REGULATOR} will hold a fitness to practise hearing about the concerns raised in the "
             "referral of 2 February 2026.")
    pdf.field("Date", "Tuesday 21 April 2026, 10:00")
    pdf.field("Place", "Online hearing (video link sent one week before)")
    pdf.field("Expected length", "2 days")
    pdf.y -= 8
    pdf.para("You may attend and be represented. Please tell the hearings team by 7 April 2026 whether you "
             "will attend and send any documents you want the panel to read.")
    pdf.text("Hearings team", 10.5, bold=True)
    pdf.text("hearings@epsc.example", 10.5)
    pdf.save()
    return buffer.getvalue()


# ---------------------------------------------------------------- scanned and photographed documents
def scanned_page(lines, size=(1240, 1754), seed=7):
    """Renders typed lines as a slightly skewed, noisy grayscale scan at 150 dpi."""
    rnd = random.Random(seed)
    img = Image.new("L", size, 250)
    draw = ImageDraw.Draw(img)
    f = font(30, TYPEWRITER)
    y = 150
    for line in lines:
        draw.text((130, y), line, fill=35, font=f)
        y += 46
    for _ in range(1600):  # scanner dust
        x, yy = rnd.randrange(size[0]), rnd.randrange(size[1])
        draw.point((x, yy), fill=rnd.randrange(120, 200))
    img = img.rotate(0.6, resample=Image.BICUBIC, fillcolor=245)
    return img.filter(ImageFilter.GaussianBlur(0.6))


def investigation_notes():
    """A PDF page that is only an image, with no text layer: the pipeline OCRs it."""
    img = scanned_page([
        f"{REGULATOR_SHORT} - SITE VISIT NOTES          Exhibit PS/02",
        "",
        "Date of visit:   9 February 2026, 10:30",
        "Location:        Riverside Community Pharmacy",
        "Investigator:    Priya Shah",
        "Present:         Chris Walker (pharmacy manager)",
        "",
        "1. Shelving: 5 mg and 50 mg tablets now on separate",
        "   shelves, warning labels in place. Photos taken.",
        "2. CD register: checked January pages. Three late",
        "   entries (12, 19 and 23 January). Balances correct.",
        "3. Rota: two staff on 27 January afternoon, planned",
        "   three. Manager confirms a vacancy since December.",
        "4. Two-person check for look-alike medicines started",
        "   on 3 February. Signed staff briefing seen.",
        "",
        "Next steps: request statement from C. Walker and",
        "the registrant's written response.",
        "",
        "P. Shah",
    ])
    jpeg = io.BytesIO()
    img.save(jpeg, "JPEG", quality=55)
    c = canvas.Canvas(str(OUT / "PS02 investigation notes.pdf"), pagesize=A4, invariant=1)
    c.setTitle("Investigation notes (scan)")
    c.setAuthor(REGULATOR)
    c.drawImage(ImageReader(io.BytesIO(jpeg.getvalue())), 0, 0, width=A4[0], height=A4[1])
    c.save()


def handwritten_note():
    """A photographed handwritten note (JPEG, no text): the pipeline converts and OCRs it."""
    rnd = random.Random(3)
    img = Image.new("L", (1240, 1754), 246)
    draw = ImageDraw.Draw(img)
    for y in range(260, 1650, 70):  # ruled paper
        draw.line((90, y + 46, 1150, y + 46), fill=205, width=2)
    f = font(44, HAND)
    lines = [
        "28 Jan 2026 - 09:20",
        "Call from the carer of the patient:",
        "box says 50 mg, prescription is 5 mg.",
        "Patient has not taken any tablets.",
        "Checked the bag record: handed out",
        "yesterday 16:40, checked by J. Ellis.",
        "Arranged home delivery of the correct",
        "tablets today before 12:00.",
        "Told J. Ellis when he arrived.",
        "Incident form to follow.",
        "",
        "C. Walker",
    ]
    y = 262
    for line in lines:
        draw.text((110 + rnd.randrange(-6, 6), y), line, fill=40, font=f)
        y += 70
    draw.text((110, 120), "Exhibit CW/01", fill=40, font=font(36, HAND))
    img = img.rotate(-0.8, resample=Image.BICUBIC, fillcolor=240)
    img.save(OUT / "CW01 handwritten note.jpg", "JPEG", quality=60, dpi=(150, 150))


# ---------------------------------------------------------------- Word documents
def new_docx(title):
    doc = Document()
    props = doc.core_properties
    props.author = REGULATOR
    props.last_modified_by = REGULATOR
    props.title = title
    props.created = props.modified = FIXED_DATE.replace(tzinfo=None)
    props.revision = 1
    section = doc.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)  # A4
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)
    return doc


def witness_statement(name, role, exhibits, paragraphs, filename):
    doc = new_docx(f"Witness statement of {name}")
    doc.add_paragraph(REGULATOR).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading(f"Witness statement of {name}", level=1)
    doc.add_paragraph(f"Case reference: {CASE_REF}")
    doc.add_paragraph(f"Registrant: {REGISTRANT} ({REGISTRATION})")
    doc.add_paragraph(f"Occupation: {role}")
    doc.add_paragraph(f"Exhibits referred to: {', '.join(exhibits)}")
    for i, text in enumerate(paragraphs, start=1):
        doc.add_paragraph(f"{i}. {text}")
    doc.add_heading("Statement of truth", level=2)
    doc.add_paragraph("I believe that the facts stated in this witness statement are true.")
    doc.add_paragraph(f"Signed: {name[0]}. {name.split()[-1]}")
    doc.add_paragraph("Date: 16 February 2026")
    doc.save(OUT / filename)


def word_documents():
    witness_statement(
        "Chris Walker", "Pharmacy manager, Riverside Community Pharmacy", ["CW/01", "CW/02"],
        [
            "I have managed Riverside Community Pharmacy since 2021. The registrant has worked there as a "
            "pharmacist since March 2024.",
            "On 28 January 2026 at about 09:20 I took a phone call from the carer of a patient who had "
            "received 50 mg tablets instead of 5 mg tablets. I wrote a note of the call at the time. This is "
            "exhibit CW/01.",
            "I arranged for the correct tablets to be delivered to the patient's home the same morning and I "
            "told the registrant about the call when he arrived at work.",
            "I completed the pharmacy's incident report on 30 January 2026. This is exhibit CW/02.",
            "We were one member of staff short on the afternoon of 27 January because of a vacancy that we "
            "had not yet filled.",
        ],
        "Witness statement - Chris Walker.docx")
    witness_statement(
        "Priya Shah", f"Investigator, {REGULATOR}", ["PS/01", "PS/02"],
        [
            "I am an investigator employed by the Council. I was asked to investigate the referral made by "
            "Mr Walker on 2 February 2026.",
            "I visited the pharmacy on 9 February 2026 and made handwritten notes during the visit. A scan of "
            "those notes is exhibit PS/02.",
            "I reviewed the documents listed in my investigation report and wrote the report on 24 February "
            "2026. This is exhibit PS/01.",
            "The registrant co-operated fully with the investigation and provided a written response through "
            "his representative.",
        ],
        "Witness statement - Priya Shah.docx")

    doc = new_docx("Meeting notes")
    doc.add_heading("Case review meeting - notes", level=1)
    doc.add_paragraph(f"Case reference: {CASE_REF}")
    doc.add_paragraph("Date: 26 February 2026, 14:00 (video call)")
    doc.add_paragraph("Attendees: Sam Turner (case officer), Priya Shah (investigator), Morgan Lee (legal adviser)")
    doc.add_heading("Discussion", level=2)
    for item in [
        "The investigation report is complete. The facts are not disputed.",
        "The registrant has completed a refresher course and the pharmacy has changed its checking process.",
        "The legal adviser recommends a hearing to decide on current impairment.",
        "The case officer will prepare the hearing bundle and the witness bundle.",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    doc.add_heading("Actions", level=2)
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    for cell, text in zip(table.rows[0].cells, ["Action", "Owner", "Due"]):
        cell.text = text
    for row in [("Send notice of hearing", "Sam Turner", "6 March 2026"),
                ("Prepare bundles", "Sam Turner", "31 March 2026"),
                ("Confirm witness availability", "Priya Shah", "13 March 2026")]:
        cells = table.add_row().cells
        for cell, text in zip(cells, row):
            cell.text = text
    doc.save(OUT / "Meeting notes.docx")

    doc = new_docx("Character reference")
    doc.add_paragraph("Northbridge Health Centre\n4 Station Lane\nNorthbridge NB2 7LP")
    doc.add_paragraph("10 March 2026")
    doc.add_paragraph(f"To the panel of the {REGULATOR}")
    doc.add_paragraph(f"Re: {REGISTRANT}, registration number {REGISTRATION}")
    for text in [
        f"I am a general practitioner and I have worked with {REGISTRANT} for two years, mostly on "
        "medication reviews for patients in care homes.",
        "In that time I have found Mr Ellis careful, honest and willing to ask for help. He told me about the "
        "dispensing error himself and explained what the pharmacy had changed as a result.",
        "I have no hesitation in continuing to work with him.",
    ]:
        doc.add_paragraph(text)
    doc.add_paragraph("Yours faithfully,\n\nDr Helen Brooks")
    doc.save(OUT / "Character reference.docx")

    # Cover sheet template: the pipeline replaces the bracketed placeholders.
    doc = new_docx("Exhibit cover sheet template")
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("EXHIBIT COVER SHEET")
    run.bold = True
    run.font.size = Pt(18)
    for text in [REGULATOR, "v", "[ Defendant name ]", "", "Case reference: [ Case reference ]", ""]:
        p = doc.add_paragraph(text)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Exhibit [ Exhibit ref ]")
    r.bold = True
    r.font.size = Pt(28)
    p = doc.add_paragraph("[ Document name ]")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("This is the exhibit marked [ Exhibit ref ] referred to in the witness statement of ")
    # Split over two runs on purpose: Word often does this, and the pipeline joins them.
    p.add_run("[ Witness ")
    p.add_run("name ]").bold = True
    p.add_run(".")
    doc.save(OUT / "Exhibit cover sheet template.docx")


def registrant_response_docx():
    doc = new_docx("Registrant response")
    doc.add_paragraph("Reid Legal LLP\n20 Bridge Street\nNorthbridge NB1 3AA")
    doc.add_paragraph("16 March 2026")
    doc.add_paragraph(f"Hearings team, {REGULATOR}")
    doc.add_paragraph(f"Re: {CASE_REF} - written response of {REGISTRANT}")
    for text in [
        "Our client admits the dispensing error of 27 January 2026 and the three late entries in the "
        "controlled drugs register.",
        "Our client reported the error to the pharmacy manager, apologised to the patient's carer and has "
        "completed a refresher course on safe dispensing.",
        "We will attend the hearing on 21 April 2026 with our client. We enclose a character reference.",
    ]:
        doc.add_paragraph(text)
    doc.add_paragraph("Yours faithfully,\n\nAlex Reid\nReid Legal LLP")
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# ---------------------------------------------------------------- emails
def email(path, sender, to, subject, sent, paragraphs, attachment_name, attachment, mime_type, message_id):
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = to
    msg["Subject"] = subject
    msg["Date"] = format_datetime(sent)
    msg["Message-ID"] = message_id
    msg.set_content("\n\n".join(paragraphs))
    msg.add_alternative("<html><body>" + "".join(f"<p>{p}</p>" for p in paragraphs).replace("\n", "<br>")
                        + "</body></html>", subtype="html")
    maintype, subtype = mime_type.split("/")
    msg.add_attachment(attachment, maintype=maintype, subtype=subtype, filename=attachment_name)
    (OUT / path).write_bytes(msg.as_bytes(policy=SMTP))  # CRLF line endings, as mail servers store them


def emails():
    email("Email - Notice of hearing.eml",
          f"{REGULATOR_SHORT} Hearings Team <hearings@epsc.example>",
          f"{REGISTRANT} <jordan.ellis@mail.example>",
          f"Notice of hearing - {CASE_REF}",
          datetime(2026, 3, 5, 10, 15, tzinfo=timezone.utc),
          ["Dear Mr Ellis,",
           "Please find attached the notice of your fitness to practise hearing on 21 April 2026.",
           "Please tell us by 7 April 2026 whether you will attend.",
           "Kind regards,\nSam Turner\nCase officer, Hearings Team"],
          "Notice of hearing.pdf", notice_of_hearing(), "application/pdf",
          "<notice-0142@epsc.example>")
    email("Email - Registrant response.eml",
          "Alex Reid <a.reid@reid-legal.example>",
          f"{REGULATOR_SHORT} Hearings Team <hearings@epsc.example>",
          f"RE: Notice of hearing - {CASE_REF}",
          datetime(2026, 3, 16, 15, 42, tzinfo=timezone.utc),
          ["Dear Sam,",
           "We act for Mr Ellis. Please find attached his written response.",
           "We confirm that Mr Ellis will attend the hearing on 21 April 2026.",
           "Best regards,\nAlex Reid\nReid Legal LLP"],
          "Registrant response.docx", registrant_response_docx(),
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          "<response-0142@reid-legal.example>")


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    referral_form()
    incident_report()
    investigation_report()
    investigation_notes()
    handwritten_note()
    word_documents()
    emails()
    for f in sorted(OUT.iterdir()):
        print(f"{f.stat().st_size:>8}  {f.relative_to(OUT.parent.parent)}")
