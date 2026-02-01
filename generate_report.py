import pandas as pd
import numpy as np
from pylatex import Document, Section, Subsection, Command, Figure, NewPage, Tabular, NoEscape
from pylatex.utils import bold, italic

# ---------------------------------------------------------
# USER INPUTS – UPDATE THESE
# ---------------------------------------------------------
EXCEL_PATH = "Book1.xlsx"        # Your Excel file path
BEAM_IMAGE_PATH = "beam.png"    # Provided beam image
OUTPUT_PDF = "Engineering_Report.pdf"

# ---------------------------------------------------------
# FUNCTION TO COMPUTE SFD/BMD FOR POINT LOADS
# ---------------------------------------------------------
def compute_sfd_bmd(length, forces):
    """
    forces = list of tuples (position, magnitude)
    """
    x = np.linspace(0, length, 400)

    # Reaction forces for simply supported beam with point loads
    R1 = sum([P * (length - a) / length for a, P in forces])
    R2 = sum([P * a / length for a, P in forces])

    shear = []
    moment = []
    for xi in x:
        V = R1
        M = R1 * xi
        for a, P in forces:
            if xi >= a:
                V -= P
                M -= P * (xi - a)
        shear.append(V)
        moment.append(M)

    return x, shear, moment, R1, R2

# ---------------------------------------------------------
# READ EXCEL FORCE DATA
# ---------------------------------------------------------
df = pd.read_excel(EXCEL_PATH)

print("\n--- DEBUG: Excel Columns ---")
print(df.columns.tolist())
print(df.head())

# Expecting columns: Position(m), Force(kN)
forces = list(zip(df["Position"], df["Force"]))

BEAM_LENGTH = df["Position"].max()

# Compute SFD & BMD
x, sfd, bmd, R1, R2 = compute_sfd_bmd(BEAM_LENGTH, forces)

# Convert coordinate arrays into TikZ pgfplots format
tikz_sfd_coords = "\n".join([f"({round(x[i],3)},{round(sfd[i],3)})" for i in range(len(x))])
tikz_bmd_coords = "\n".join([f"({round(x[i],3)},{round(bmd[i],3)})" for i in range(len(x))])

# ---------------------------------------------------------
# START PDF DOCUMENT
# ---------------------------------------------------------
doc = Document(documentclass="report")

# Title Page
doc.preamble.append(Command('title', 'Engineering Report: Beam Analysis'))
doc.preamble.append(Command('author', 'Generated Using Python & PyLaTeX'))
doc.preamble.append(Command('date', NoEscape(r'\today')))
doc.append(NoEscape(r'\maketitle'))
doc.append(NewPage())

# Table of Contents
doc.append(Command('tableofcontents'))
doc.append(NewPage())

# ---------------------------------------------------------
# INTRODUCTION
# ---------------------------------------------------------
with doc.create(Section("Introduction")):
    doc.append("This report presents an analysis of a simply supported beam subjected to given loading conditions.")
    
    with doc.create(Subsection("Beam Description")):
        with doc.create(Figure(position='h!')) as fig:
            fig.add_image(BEAM_IMAGE_PATH, width=NoEscape(r"0.8\textwidth"))
            fig.add_caption("Simply Supported Beam")

    with doc.create(Subsection("Data Source")):
        doc.append("The loading data was read from the provided Excel file using pandas.")

# ---------------------------------------------------------
# FORCE TABLE (LATEX TABULAR)
# ---------------------------------------------------------
with doc.create(Section("Input Data: Force Table")):
    with doc.create(Tabular('|c|c|')) as table:
        table.add_hline()
        table.add_row((bold("Position (m)"), bold("Force (kN)")))
        table.add_hline()
        for index, row in df.iterrows():
            table.add_row((row["Position"], row["Force"]))
            table.add_hline()

# ---------------------------------------------------------
# ANALYSIS SECTION WITH TIKZ
# ---------------------------------------------------------
doc.append(NewPage())
with doc.create(Section("Analysis")):

    # SFD
    with doc.create(Subsection("Shear Force Diagram")):
        tikz_sfd = NoEscape(r"""
\begin{tikzpicture}
\begin{axis}[
    width=14cm,
    height=6cm,
    grid=both,
    xlabel={Position (m)},
    ylabel={Shear Force (kN)},
    thick,
]
\addplot[smooth, blue] coordinates {
""" + tikz_sfd_coords + r"""
};
\end{axis}
\end{tikzpicture}
        """)
        doc.append(tikz_sfd)

    # BMD
    with doc.create(Subsection("Bending Moment Diagram")):
        tikz_bmd = NoEscape(r"""
\begin{tikzpicture}
\begin{axis}[
    width=14cm,
    height=6cm,
    grid=both,
    xlabel={Position (m)},
    ylabel={Bending Moment (kNm)},
    thick,
]
\addplot[smooth, red] coordinates {
""" + tikz_bmd_coords + r"""
};
\end{axis}
\end{tikzpicture}
        """)
        doc.append(tikz_bmd)

# ---------------------------------------------------------
# SUMMARY
# ---------------------------------------------------------
with doc.create(Section("Summary")):
    doc.append(
        "A Shear Force Diagram (SFD) shows how internal shear force varies along the beam.\n"
        "A Bending Moment Diagram (BMD) shows how internal bending varies along the beam.\n"
    )

# ---------------------------------------------------------
# GENERATE PDF
# ---------------------------------------------------------

doc.generate_pdf(
    OUTPUT_PDF,
    compiler='C:/Users/Aditya jain/AppData/Local/Programs/MiKTeX/miktex/bin/x64/pdflatex.exe',
    clean_tex=False
)
print("PDF Generated Successfully!")