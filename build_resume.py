"""Build the ATS-friendly resume.

Usage:
  python3 build_resume.py Harrison-Pielke-Lombardo-Resume.docx
  python3 build_resume.py Harrison-Pielke-Lombardo-Resume-Contract.docx contract
Then convert to PDF: soffice --headless --convert-to pdf <file>.docx
Requires python-docx.
"""
import sys
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = sys.argv[1]
FONT = "Calibri"
ACCENT = RGBColor(0x1F, 0x3A, 0x5F)

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Inches(8.5), Inches(11)
sec.left_margin = sec.right_margin = Inches(0.6)
sec.top_margin = sec.bottom_margin = Inches(0.5)
RIGHT = sec.page_width - sec.left_margin - sec.right_margin

st = doc.styles["Normal"]
st.font.name = FONT
st.font.size = Pt(10)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
st.paragraph_format.space_after = Pt(0)
st.paragraph_format.space_before = Pt(0)
st.paragraph_format.line_spacing = 1.0

props = doc.core_properties
props.title = "Harrison Pielke-Lombardo - Resume"
props.author = "Harrison Pielke-Lombardo"


def para(space_before=0, space_after=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    return p


def run(p, text, bold=False, italic=False, size=None, color=None):
    r = p.add_run(text)
    r.bold, r.italic = bold, italic
    if size:
        r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    return r


def bottom_border(p):
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", "6"), ("space", "1"), ("color", "1F3A5F")):
        b.set(qn(f"w:{k}"), v)
    bdr.append(b)
    pPr.append(bdr)


def heading(text):
    p = para(6, 2)
    run(p, text.upper(), bold=True, size=11, color=ACCENT)
    bottom_border(p)


def line_with_date(left, right, bold=True, italic=False, space_before=0, sep="  |  "):
    # Inline, no tab stops: right-aligned tabs make PDF parsers detach the date from its line.
    p = para(space_before, 0)
    run(p, left, bold=bold, italic=italic)
    run(p, sep + right, bold=False, italic=italic)
    return p


def bullet(text):
    p = para(0, 1)
    p.paragraph_format.left_indent = Inches(0.18)
    p.paragraph_format.first_line_indent = Inches(-0.13)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(0.18))
    run(p, "\u2022\t" + text)


def labeled(label, text):
    p = para(0, 1.5)
    run(p, label + ": ", bold=True)
    run(p, text)


CONTRACT = len(sys.argv) > 2 and sys.argv[2] == "contract"

# Header
p = para()
run(p, "Harrison Pielke-Lombardo", bold=True, size=20)
p = para()
if CONTRACT:
    run(p, "Senior Backend & Full-Stack Engineer  |  Available for Contract & 1099 Work", bold=True, size=12,
        color=ACCENT)
else:
    run(p, "Senior Backend & Full-Stack Software Engineer", bold=True, size=12, color=ACCENT)
p = para(2, 0)
run(p, "Denver, CO (remote or hybrid)  |  720-209-6249  |  pielkekid@gmail.com  |  "
       "linkedin.com/in/tuh8888  |  github.com/tuh8888", size=10)

heading("Summary")
p = para()
if CONTRACT:
    run(p, "Senior backend and full-stack engineer available for contract and 1099 engagements, remote or "
           "on-site in Colorado. 10 years of building production systems in Python, Clojure, TypeScript, and "
           "Java. Most recently built the document-processing pipeline, rules engine, and AWS infrastructure "
           "behind a B2B SaaS logistics platform handling 1M+ documents per month, which grew from zero to "
           "$2.1M ARR and was acquired. Comfortable owning a project end to end, working directly with "
           "clients, and getting productive quickly in an existing codebase.")
else:
    run(p, "Senior software engineer with 10 years of experience building production backend systems, "
           "data pipelines, and web applications in Python, Clojure, TypeScript, and Java. Most recently a "
           "core engineer on a 4-6 person team at a B2B SaaS logistics startup, building the pipeline that "
           "processed 1M+ documents per month and helping grow the product from zero to $2.1M ARR through "
           "acquisition. Owns features end to end, from design with product through on-call, and has "
           "mentored, interviewed, and served as technical contact for clients.")

if CONTRACT:
    heading("What I Can Deliver")
    labeled("Backend APIs and data pipelines", "REST and GraphQL services, high-volume document and data "
                                               "processing, rules engines")
    labeled("Database migrations", "zero-downtime production migrations, including Datomic to PostgreSQL "
                                   "at Datomic's size limits")
    labeled("AWS infrastructure", "Terraform-managed, event-driven systems on Lambda, ECS, SQS, SNS, and RDS")
    labeled("LLM and ML automation", "production document processing and verification that replaces "
                                     "manual review work")
    labeled("Clojure and ClojureScript", "new builds, maintenance, and modernization of existing Clojure "
                                         "systems")

heading("Skills")
labeled("Languages", "Python, Clojure, ClojureScript, TypeScript, Java, JavaScript, SQL, Bash")
if not CONTRACT:
    labeled("Backend", "REST APIs, GraphQL, data pipelines, rules engines, event-driven architecture, "
                       "database migrations")
labeled("Frontend", "TypeScript, ClojureScript, React (Reagent / re-frame), JavaScript, data visualization")
labeled("Data", "PostgreSQL, Datomic, DynamoDB, Redis, SQL, knowledge graphs (RDF)")
labeled("Cloud & DevOps", "AWS (Lambda, ECS, EC2, SQS, SNS, RDS, DynamoDB, CloudFront), Terraform, "
                          "Docker, CI/CD, Git, GitHub")
labeled("AI & ML", "LLM integration, PyTorch, NLP, AI-assisted development (Claude, Copilot)")

heading("Experience")
line_with_date("Senior Software Engineer", "Sep 2022 - Mar 2026")
line_with_date("Peruse Technology Inc. (acquired)", "Remote", sep=", ", bold=False, italic=True)
bullet("Helped scale a B2B SaaS logistics platform from zero to $2.1M ARR and 70+ business customers as "
       "part of a 4-6 person engineering team, leading to the company's acquisition.")
bullet("Designed and built the core load processing pipeline and a custom rules engine in Clojure and "
       "Python, processing 1M+ documents per month through REST and GraphQL APIs.")
bullet("Trained and deployed LLM-based models in PyTorch to automate document processing and freight broker "
       "verification, replacing work that took customers entire back-office audit teams.")
bullet("Planned and executed the zero-downtime migration of the production database from Datomic to "
       "PostgreSQL as it approached Datomic's size limits.")
bullet("Shared ownership of event-driven AWS infrastructure as code in Terraform (Lambda, ECS, SQS, SNS, RDS); "
       "took on-call rotations and contributed to the TypeScript frontend.")
if not CONTRACT:
    bullet("Owned features end to end, from design with the product team to release; mentored an engineer, "
           "interviewed candidates, and reviewed code.")
bullet("Served as technical contact on client calls and owned relationships with several client accounts.")

line_with_date("Research Software Engineer (Graduate Researcher)", "Aug 2016 - Aug 2022", space_before=4)
line_with_date("University of Colorado Anschutz Medical Campus", "Aurora, CO", sep=", ", bold=False, italic=True)
bullet("Led software for interconnected research projects across 4+ cross-functional teams, shipping "
       "5 open-source package releases, 4 dataset releases, and 6 peer-reviewed publications.")
bullet("Built Knowtator, a Java and ClojureScript / re-frame web app for text annotation and ontology "
       "development (rewritten from a Java Swing desktop app), used to create 100k+ annotations by "
       "clinicians and researchers.")
if not CONTRACT:
    bullet("Built data integration pipelines that merged large knowledge graphs across RDF, Datomic, and SQL "
           "stores for drug-target identification and diagnostic tools.")
bullet("Improved relation-extraction precision by 10% with a novel bootstrapping algorithm accelerated "
       "with GPU matrix operations.")

heading("Open Source & Speaking")
bullet("datascript-shortest-path, a graph library with a Fibonacci heap built on immutable data structures; "
       "contributor to the re-frame ecosystem (day8/re-frame-undo).")
bullet("Invited speaker, \"AI in the Transportation Sector: Challenges and Successes,\" Institute for Data, "
       "Econometrics, Algorithms, and Learning (IDEAL), Chicago, June 2024.")

heading("Education")
line_with_date("M.S., Computational Bioscience, University of Colorado Anschutz Medical Campus", "2022",
               bold=False)
line_with_date("B.S., Applied Mathematics (minors: Computer Science, Biochemistry), University of Colorado "
               "Boulder", "2016", bold=False)

doc.save(OUT)
