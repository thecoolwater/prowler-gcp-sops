"""
Generate two simple, plain DOCX SOP documents for GCP Security Audit.
- Document 1: Customer-facing GCP Prerequisites (No mentions of Prowler, generic security-audit-sa, no simulated screenshots)
- Document 2: Internal-facing Prowler Audit Execution Guide (Retains Prowler instructions, no simulated screenshots)
"""

from docx import Document
from docx.shared import Inches, Pt
import datetime

TODAY = datetime.date.today().strftime("%B %d, %Y")

def add_simple_header(doc, title, doc_id):
    """Add a simple title and document ID at the top of the first page."""
    p_title = doc.add_paragraph()
    r_title = p_title.add_run(title)
    r_title.bold = True
    r_title.font.size = Pt(18)
    r_title.font.name = "Arial"
    
    p_meta = doc.add_paragraph()
    r_meta = p_meta.add_run(f"Document ID: {doc_id}  |  Date: {TODAY}  |  Version: 1.0\n")
    r_meta.font.size = Pt(9.5)
    r_meta.italic = True
    r_meta.font.name = "Arial"

def add_simple_heading(doc, text, level):
    """Add plain Arial headings."""
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.name = "Arial"
        run.bold = True
        if level == 1:
            run.font.size = Pt(14)
        elif level == 2:
            run.font.size = Pt(12)
        else:
            run.font.size = Pt(11)

def add_simple_paragraph(doc, text, bold=False, italic=False, space_after=6):
    """Add simple Arial body text."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(10)
    run.bold = bold
    run.italic = italic
    return p

def add_simple_bullet(doc, text):
    """Add plain bullet points."""
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(3)
    p.clear()
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(10)

def add_simple_code(doc, code_text):
    """Add a plain, unshaded code block without forced line breaks to avoid wrapping issues."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(code_text)
    run.font.name = "Courier New"
    run.font.size = Pt(9)

def add_simple_table(doc, headers, rows):
    """Add a simple unstyled table."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.autofit = True
    
    # Header
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(header)
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(9.5)
        
    # Rows
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_text in enumerate(row_data):
            cell = table.rows[row_idx + 1].cells[col_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(str(cell_text))
            run.font.name = "Arial"
            run.font.size = Pt(9)
            
    doc.add_paragraph() # space after table

def add_simple_image(doc, img_path, caption):
    """Add an image with a simple caption centered."""
    try:
        p_img = doc.add_paragraph()
        p_img.alignment = 1 # Center
        p_img.paragraph_format.space_before = Pt(6)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.add_run().add_picture(img_path, width=Inches(5.0))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = 1 # Center
        p_cap.paragraph_format.space_after = Pt(8)
        run = p_cap.add_run(f"Figure: {caption}")
        run.font.name = "Arial"
        run.font.size = Pt(8.5)
        run.italic = True
    except Exception as e:
        print(f"Warning: Could not add image {img_path} - {e}")

# ═══════════════════════════════════════════════════════════════════════════════
# GENERATE DOCUMENT 1: PREREQUISITES (CUSTOMER-FACING)
# - No mentions of Prowler.
# - Generic audit/assessment wording.
# - Generic service account names.
# - No simulated screenshots.
# ═══════════════════════════════════════════════════════════════════════════════
def make_simple_prereqs():
    doc = Document()
    DOC_ID = "SOP-GCP-SECURITY-ASSESSMENT-PREREQS"
    
    add_simple_header(doc, "GCP Prerequisites Setup for Cloud Security Assessment", DOC_ID)
    
    add_simple_heading(doc, "1. Introduction", level=1)
    add_simple_paragraph(doc, "This document outlines the steps required to prepare your Google Cloud Platform (GCP) environment for a third-party read-only cloud security posture assessment. All steps are performed on the customer side prior to the audit.")
    
    add_simple_heading(doc, "2. Requirements", level=1)
    add_simple_bullet(doc, "A GCP account with administrative access to the target projects.")
    add_simple_bullet(doc, "Google Cloud SDK (gcloud CLI) installed and configured on your machine, or access to GCP Cloud Shell.")
    
    add_simple_heading(doc, "3. Enable Required GCP APIs", level=1)
    add_simple_paragraph(doc, "You must enable the Identity and Access Management (IAM) API in your target project to allow resource discovery.")
    add_simple_paragraph(doc, "You can enable this in the Google Cloud Console (APIs & Services > Library > Search for IAM API) or run the following command:")
    add_simple_code(doc, "gcloud services enable iam.googleapis.com --project <YOUR_PROJECT_ID>")
    
    add_simple_paragraph(doc, "It is also recommended to enable the Cloud Resource Manager API:")
    add_simple_code(doc, "gcloud services enable cloudresourcemanager.googleapis.com --project <YOUR_PROJECT_ID>")
    
    add_simple_heading(doc, "4. Create a Service Account", level=1)
    add_simple_paragraph(doc, "Create a dedicated service account that the assessment team will use to access your project resource configurations.")
    add_simple_paragraph(doc, "You can do this by navigating to IAM & Admin > Service Accounts in the GCP Console.")
    add_simple_image(doc, "service-account-page.png", "Service Accounts Page in Google Cloud Console")
    
    add_simple_paragraph(doc, "Run the following command to create the service account, or use the '+ CREATE SERVICE ACCOUNT' button in the Console:")
    add_simple_code(doc, "gcloud iam service-accounts create security-audit-sa --display-name=\"Security Audit Service Account\" --project=<YOUR_PROJECT_ID>")
    add_simple_image(doc, "create-service-account.png", "Creating a Service Account in the GCP Console")
    
    add_simple_heading(doc, "5. Assign IAM Permissions", level=1)
    add_simple_paragraph(doc, "The assessment team requires read-only permissions. Assign the standard 'Viewer' and 'Service Usage Consumer' roles to the service account.")
    add_simple_paragraph(doc, "Run these commands to bind the roles to the service account, or assign them under the 'Grant access' dialog in IAM:")
    add_simple_code(doc, 
        "gcloud projects add-iam-policy-binding <YOUR_PROJECT_ID> --member=\"serviceAccount:security-audit-sa@<YOUR_PROJECT_ID>.iam.gserviceaccount.com\" --role=\"roles/viewer\"\n\n"
        "gcloud projects add-iam-policy-binding <YOUR_PROJECT_ID> --member=\"serviceAccount:security-audit-sa@<YOUR_PROJECT_ID>.iam.gserviceaccount.com\" --role=\"roles/serviceusage.serviceUsageConsumer\""
    )
    add_simple_image(doc, "service-account-permissions.png", "Assigning IAM Roles and Permissions in GCP")
    
    add_simple_paragraph(doc, "Important: Replace <YOUR_PROJECT_ID> in all commands with your actual GCP Project ID.", bold=True)

    add_simple_heading(doc, "6. Generate JSON Key Credentials", level=1)
    add_simple_paragraph(doc, "You must generate and download a JSON key file for the service account to provide to the assessment team.")
    add_simple_paragraph(doc, "Navigate to the Keys tab under the service account details, click 'ADD KEY', and select 'Create new key':")
    add_simple_image(doc, "create-new-key.png", "Adding a Key to the Service Account")
    
    add_simple_paragraph(doc, "Select JSON as the format and click 'CREATE':")
    add_simple_image(doc, "json-key.png", "Selecting the JSON Key Format")
    
    add_simple_paragraph(doc, "Alternatively, run the following command to generate the key:")
    add_simple_code(doc, 
        "gcloud iam service-accounts keys create security-audit-key.json --iam-account=security-audit-sa@<YOUR_PROJECT_ID>.iam.gserviceaccount.com"
    )
    add_simple_paragraph(doc, "This downloads a file named 'security-audit-key.json' to your current working directory. Share this file securely with the security assessment team.")

    doc.save("SOP-GCP-PREREQUISITES-CUSTOMER-v1.0.docx")

# ═══════════════════════════════════════════════════════════════════════════════
# GENERATE DOCUMENT 2: AUDIT EXECUTION (INTERNAL-ONLY)
# - Retains Prowler references and Prowler audit steps.
# - No simulated screenshots.
# ═══════════════════════════════════════════════════════════════════════════════
def make_simple_audit():
    doc = Document()
    DOC_ID = "SOP-GCP-PROWLER-AUDIT"
    
    add_simple_header(doc, "GCP Prowler Audit Execution Guide (Internal Use)", DOC_ID)
    
    add_simple_heading(doc, "1. Introduction", level=1)
    add_simple_paragraph(doc, "This guide shows how to run security assessments on GCP using Prowler, configure filters, and view reports.")
    add_simple_paragraph(doc, "Use the credentials file provided by the customer (e.g. security-audit-key.json).")
    
    add_simple_heading(doc, "2. Configure Authentication", level=1)
    add_simple_paragraph(doc, "Point Prowler to the customer's shared credentials key file:")
    
    add_simple_paragraph(doc, "On Linux/macOS:")
    add_simple_code(doc, 
        "export GOOGLE_APPLICATION_CREDENTIALS=\"/absolute/path/to/security-audit-key.json\"\n"
        "export GOOGLE_CLOUD_QUOTA_PROJECT=\"<CUSTOMER_PROJECT_ID>\""
    )
    
    add_simple_paragraph(doc, "On Windows (PowerShell):")
    add_simple_code(doc, 
        "$env:GOOGLE_APPLICATION_CREDENTIALS = \"C:\\path\\to\\security-audit-key.json\"\n"
        "$env:GOOGLE_CLOUD_QUOTA_PROJECT = \"<CUSTOMER_PROJECT_ID>\""
    )
    
    add_simple_heading(doc, "3. Running the Audit", level=1)
    add_simple_paragraph(doc, "With your environment variables set, execute Prowler using the following command:")
    add_simple_code(doc, "prowler gcp")
    
    add_simple_heading(doc, "4. Common Filter Options", level=1)
    add_simple_paragraph(doc, "You can customize the audit using these command flags:")
    
    add_simple_paragraph(doc, "Scan specific projects:", bold=True)
    add_simple_code(doc, "prowler gcp --project-ids <CUSTOMER_PROJECT_ID>")
    
    add_simple_paragraph(doc, "Scan specific services only (e.g. storage and iam):", bold=True)
    add_simple_code(doc, "prowler gcp --services storage iam")
    
    add_simple_paragraph(doc, "Filter by severity (e.g. only show critical and high findings):", bold=True)
    add_simple_code(doc, "prowler gcp --severity critical high")

    add_simple_heading(doc, "5. Running Compliance Scans", level=1)
    add_simple_paragraph(doc, "To check compliance against standard security frameworks, use the --compliance flag:")
    
    add_simple_paragraph(doc, "CIS GCP Foundations Benchmark:", bold=True)
    add_simple_code(doc, "prowler gcp --compliance cis_2.0_gcp")
    
    add_simple_paragraph(doc, "Other available frameworks:", bold=True)
    add_simple_table(doc, 
        ["Framework Name", "Prowler Identifier"],
        [
            ["CIS GCP Foundations v3.0", "cis_3.0_gcp"],
            ["NIST 800-53", "nist_800_53_revision_5_gcp"],
            ["PCI-DSS v4.0", "pci_4.0_gcp"],
            ["SOC 2", "soc2_gcp"],
            ["ISO 27001:2022", "iso27001_2022_gcp"],
        ]
    )

    add_simple_heading(doc, "6. Outputs and Reports", level=1)
    add_simple_paragraph(doc, "Prowler automatically saves the scan results in a directory called 'output' relative to where you ran the command.")
    add_simple_paragraph(doc, "It generates three file formats by default:")
    add_simple_bullet(doc, ".html - Human-readable report dashboard")
    add_simple_bullet(doc, ".csv - Raw data for spreadsheets")
    add_simple_bullet(doc, ".json - Structured data for APIs or log managers")
    
    add_simple_paragraph(doc, "To output reports to a specific folder, use the -o flag:")
    add_simple_code(doc, "prowler gcp -o ./my-gcp-audit-reports")

    add_simple_heading(doc, "7. Troubleshooting", level=1)
    add_simple_bullet(doc, "DefaultCredentialsError: Ensure GOOGLE_APPLICATION_CREDENTIALS points to the correct location of your JSON key file.")
    add_simple_bullet(doc, "403 Forbidden: Ensure the service account has been granted roles/viewer and roles/serviceusage.serviceUsageConsumer in the project.")
    add_simple_bullet(doc, "API not enabled: Enable the missing API using Cloud Console or gcloud cli.")

    doc.save("SOP-GCP-AUDIT-INTERNAL-v1.0.docx")

if __name__ == "__main__":
    make_simple_prereqs()
    make_simple_audit()
    print("Successfully generated simple documents.")
