# import streamlit as st
# import pandas as pd
# from io import BytesIO, StringIO
# import fitz  # PyMuPDF
# from pathlib import Path
# from docx import Document
# from docx.shared import Inches
# from docx.enum.text import WD_ALIGN_PARAGRAPH
# import tempfile
# from docxtpl import DocxTemplate


# st.set_page_config(page_title="Guna Solar OPEx Proposal Generator", layout="centered")
# st.title("OPEx Proposal Generator")

# # Helper to replace unsupported characters for Word output
# def safe_text(text: str) -> str:
#     return text.replace("₂", "2")

# # === User inputs ===
# from datetime import date as dt_date

# proposal_no = st.text_input(
#     "Proposal Number",
#     value="GSPL/PPA/2025-26/026R1"
# )

# proposal_date = st.date_input(
#     "Date",
#     value=dt_date(2025, 12, 6)
# )

# off_taker = st.text_input(
#     "Off Taker",
#     value="Anand Ranganathan"
# )

# designation = st.text_input(
#     "Designation",
#     value="Chairman"
# )

# company_name = st.text_input(
#     "Enter Site Name",
#     value="ABC Pvt Ltd"
# )

# capacity_plant = st.number_input(
#     "System Capacity (kWp)",
#     min_value=1,
#     value=100
# )

# short_address = st.text_input(
#     "Location",
#     value="Irungatukottai"
# )

# long_address = st.text_input(
#     "Complete Address",
#     value="G-26, Katrambakkam Road, Sriperumbudur, Malayambakkam"
# )

# deposit = st.number_input(
#     "Deposit Amount (INR in Lakhs)",
#     min_value=0.0,
#     value=50.0
# )

# tariff = st.number_input(
#     "Tariff for the First year",
#     value = 4.0
# )

# increment = st.number_input(
#     "Annual Increment (%)",
#     value = 1.0
# )

# annual_generation = st.number_input(
#     'Annual Generation in Lakhs',
#     value = 10.00
# )

# P_nom_kWp_str = st.text_input("Capacity (power)", value='9.4')
# try:
#     capacity_plant = float(capacity_plant)
# except ValueError:
#     st.error("Please enter a valid number for Capacity (power).")
#     st.stop()

# # template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])
# # csv_file = st.file_uploader("Upload PVsyst Hourly CSV", type=["csv"])
# # pdf_file = st.file_uploader("Upload PVsyst PDF Report", type=["pdf"])

# # if template_file and csv_file and pdf_file:
#     # --- Read CSV and find header line ---
#     # csv_bytes = csv_file.read()
#     # csv_str = csv_bytes.decode("cp1252").splitlines()

#     header_line = None
#     for i, line in enumerate(csv_str):
#         if line.lower().startswith("date"):
#             header_line = i
#             break

#     if header_line is None:
#         st.error("Could not find header row in CSV.")
#     else:
#         df = pd.read_csv(
#             StringIO("\n".join(csv_str)),
#             skiprows=header_line,
#             header=0,
#             encoding="cp1252"
#         )

#         # Drop units row
#         df = df.drop(index=0).reset_index(drop=True)
#         df = df.dropna()

#         # --- KPI calculations ---
#         df["E_Grid"] = pd.to_numeric(df["E_Grid"], errors="coerce")
#         annual_energy_mwh = df["E_Grid"].sum() / 1_000_000

#         df["PR"] = pd.to_numeric(df["PR"], errors="coerce")
#         pr_percent = df.loc[df["PR"] != 0, "PR"].mean() * 100

#         #P_nom_kWp = 9.4  # system size
#         specific_yield = (annual_energy_mwh * 1000) / capacity_plant

#         emission_factor = 0.82
#         co2_savings_tons = annual_energy_mwh * emission_factor

#         # --- Show results ---
#         st.subheader("KPI Summary")
#         st.write(f"**Annual Energy:** {annual_energy_mwh:.2f} MWh")
#         st.write(f"**Performance Ratio:** {pr_percent:.2f} %")
#         st.write(f"**Specific Yield:** {specific_yield:.1f} kWh/kWp")
#         st.write(f"**CO₂ Savings:** {co2_savings_tons:.1f} tons/year")

#         # --- Extract Loss Diagram from PDF ---
#         pdf_bytes = pdf_file.read()
#         pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")

#         try:
#             page_index = 4  # adjust if needed
#             rect = fitz.Rect(50, 122, 550, 600)
#             pix = pdf_doc[page_index].get_pixmap(clip=rect, dpi=200)

#             img_bytes = BytesIO(pix.tobytes("png"))
#             st.subheader("Loss Diagram")
#             st.image(img_bytes, caption="Extracted Loss Diagram", use_column_width=True)

#             # Save loss diagram to a temp file for docx insertion
#             with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
#                 tmp_img.write(img_bytes.getbuffer())
#                 tmp_img_path = Path(tmp_img.name)

#             # --- Generate DOCX ---
#             if st.button("Generate Proposal DOCX"):
#                 # Load template from uploaded file
#                 with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_tpl:
#                     tmp_tpl.write(template_file.read())
#                     tmp_tpl_path = Path(tmp_tpl.name)

#                 doc = Document(tmp_tpl_path)

#                 # Placeholder text mapping
#                 mapping = {
#                     "{{ClientName}}": company_name,
#                     "{{Site}}": "Navalur",
#                     "{{Capacity}}": f"{capacity_plant} kW",
#                     "{{Energy}}": f"{annual_energy_mwh:.2f} MWh",
#                     "{{PR}}": f"{pr_percent:.2f} %",
#                     "{{SpecificYield}}": f"{specific_yield:.1f} kWh/kWp",
#                     "{{CO2Savings}}": f"{co2_savings_tons:.1f} tons/year"
#                 }

#                 # Image placeholders
#                 images = {
#                     "{{LossDiagram}}": tmp_img_path
#                 }

#                 # Replace text in paragraphs
#                 for p in doc.paragraphs:
#                     for key, val in mapping.items():
#                         if key in p.text:
#                             p.text = p.text.replace(key, safe_text(str(val)))

#                 # Replace text in tables
#                 for table in doc.tables:
#                     for row in table.rows:
#                         for cell in row.cells:
#                             for key, val in mapping.items():
#                                 if key in cell.text:
#                                     cell.text = cell.text.replace(key, safe_text(str(val)))

#                 # Replace image placeholders
#                 for p in doc.paragraphs:
#                     for key, img_path in images.items():
#                         if key in p.text:
#                             p.text = ""
#                             run = p.add_run()
#                             run.add_picture(str(img_path), width=Inches(4))
#                             p.alignment = WD_ALIGN_PARAGRAPH.CENTER

#                 # Save to BytesIO for download
#                 output_docx = BytesIO()
#                 doc.save(output_docx)
#                 output_docx.seek(0)

#                 st.download_button(
#                     label="⬇️ Download Proposal DOCX",
#                     data=output_docx,
#                     file_name="FilledProposal.docx",
#                     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
#                 )

#         except Exception as e:
#             st.error(f"Could not extract loss diagram: {e}")


import streamlit as st
# import pandas as pd
from io import BytesIO
# import fitz  # PyMuPDF
from pathlib import Path
# from docx import Document
# from docx.shared import Inches
# from docx.enum.text import WD_ALIGN_PARAGRAPH
# import tempfile
from docxtpl import DocxTemplate


st.set_page_config(page_title="Guna Solar OPEx Proposal Generator", layout="centered")
st.title("OPEx Proposal Generator")

# Helper to replace unsupported characters for Word output


proposal_no = st.text_input(
    "Proposal Number",
    placeholder="GSPL/PPA/2025-26/026R1"
)

proposal_date = st.date_input(
    "Date",
    help="Select proposal date (e.g. 6 Dec 2025)"
)


off_taker = st.text_input(
    "Off Taker",
    placeholder="Anand Ranganathan"
)

designation = st.text_input(
    "Designation",
    placeholder="Chairman"
)

company_name = st.text_input(
    "Enter Site Name",
    placeholder="ABC Pvt Ltd"
)

capacity_plant = st.number_input(
    "System Capacity (kWp)",
    #min_value=1,
    help="Example: 100"
)

short_address = st.text_input(
    "Location",
    placeholder="Irungatukottai"
)

long_address = st.text_input(
    "Complete Address",
    placeholder="G-26, Katrambakkam Road, Sriperumbudur, Malayambakkam"
)

deposit = st.number_input(
    "Deposit Amount (INR in Lakhs)",
    min_value=0.0,
    help="Example: 100"
)

tariff = st.number_input(
    "Tariff for the First year",
    help="Example: 4.0"
)

increment = st.number_input(
    "Annual Increment (%)",
    help = "Example: 1.0"
)

annual_generation = st.number_input(
    'Annual Generation in Lakhs',
    help= "Example: 10.0"
)
def smart_number(x):
    if float(x).is_integer():
        return str(int(x))
    return str(x)   

context = {
    "proposal_no": proposal_no,
    "proposal_date": proposal_date.strftime("%d %B %Y"),
    "off_taker": off_taker,
    "designation": designation,
    "company_name": company_name,
    "capacity_plant": smart_number(capacity_plant),
    "short_address": short_address,
    "long_address": long_address,
    "deposit": smart_number(deposit),
    "tariff": smart_number(tariff),
    "increment": smart_number(increment),
    "annual_generation": smart_number(annual_generation),
}


if st.button("Generate Proposal"):
    
    if not proposal_no or not company_name or not off_taker:
        st.error("Please fill all mandatory fields.")
    else:

        template_path = Path ("templates/opex_proposal_template.docx")

        doc = DocxTemplate(template_path)
        doc.render(context)

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        file_name = f"OPEx_Proposal_{company_name.replace(' ','_')}.docx"

        st.success("Proposal generated successfully!")

        st.download_button(
            label = "Download Proposal",
            data = buffer,
            file_name = file_name,
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )