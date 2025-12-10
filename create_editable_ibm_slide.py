import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# --- CONFIGURATION ---
SLIDE_WIDTH = Inches(24)  # Wider canvas to allow perfect spacing
SLIDE_HEIGHT = Inches(13.5)
FONT_NAME = "Arial"

# --- COLORS (Color-picked from reference image) ---
COLORS = {
    # Headers
    "HDR_BLUE": RGBColor(22, 54, 92),      # Deep Blue (Security, AI, Data)
    "HDR_PURPLE": RGBColor(112, 48, 160),  # Purple (App Dev, Auto, Asset Mgmt)
    "HDR_GREY": RGBColor(64, 64, 64),      # Dark Grey (Client Apps, Infrastructure)
    "HDR_RED": RGBColor(255, 0, 0),        # Red (OpenShift text)
    
    # Legend Status Colors
    "STS_ENTITLED": RGBColor(153, 204, 255), # Light Blue
    "STS_OPP": RGBColor(146, 208, 80),       # Light Green
    "STS_EXPLORE": RGBColor(255, 230, 153),  # Light Yellow
    "STS_RISK": RGBColor(244, 176, 132),     # Light Red/Orange
    
    # UI Elements
    "BTN_BLUE": RGBColor(68, 114, 196),     # Upload Button
    "BG_HEADER": RGBColor(255, 255, 255),   # White top header
    "BOX_FILL": RGBColor(255, 255, 255),    # White product boxes
    "TEXT_STD": RGBColor(0, 0, 0),          # Black text
    "TEXT_WHITE": RGBColor(255, 255, 255),  # White text
    "BORDER_GREY": RGBColor(180, 180, 180)  # Subtle borders
}

# --- DATA STRUCTURE ---
# Organized exactly by the visual columns in the image

# 1. LEFT COLUMN
LEFT_COL = [
    {"title": "Security", "color": COLORS["HDR_BLUE"], "products": [
        "Data Security", # Sub-header style in box
        "Guardium Data Encryption", "Guardium Data Protection", "Guardium Data Security Center", 
        "Guardium Discover and Classify", "Guardium Key Lifecycle Management",
        "Identity & Access Mgmt", # Sub-header style
        "HashiCorp Boundary", "HashiCorp Consul", "HashiCorp Vault", "ILMT", 
        "Security Verify (IAM)", "Security MaaS 360", "Trusteer (Anti-fraud)"
    ]}
]

# 2. CENTER STACK
#    Row 1: Client Apps
CENTER_ROW_1 = {
    "title": "Client Applications", "color": COLORS["HDR_GREY"], 
    "products": ["ERP", "CRM", "B2B", "B2C", "B2E", "Omnichannel", "CRM (on-prem)", "IA",
                 "Fraud", "Credit", "PCP", "Supply Chain", "Engineering / Network", "Portal / Mobile / APP", "Payment Instantaneous", "Customer Service"]
}

#    Row 2: The 6 Pillars
CENTER_ROW_2 = [
    {"title": "AI Assistants", "color": COLORS["HDR_BLUE"], "products": ["Automation", "Blueworks Live", "Business Analytics", "Business Automation", "CP4BA", "Cognos Analytics", "Decision Mgmt", "Planning Analytics", "Process Mining", "RPA", "SPSS Modeler", "watsonx Assistants", "watsonx BI Assistant", "watsonx Code Assistant", "watsonx Orchestrate", "Workflow Automation"]},
    {"title": "AI/MLOps", "color": COLORS["HDR_BLUE"], "products": ["CP4D", "OpenPages", "Orchestrate (SaaS)", "WCA Ansible & Java", "WCAz", "watsonx.ai", "watsonx.governance"]},
    {"title": "Databases", "color": COLORS["HDR_BLUE"], "products": ["CM8", "CMOD", "CP4D", "Capture", "Cloudera", "Content", "DB2", "Database Eco", "FileNet", "Hadoop", "Informix", "Netezza", "watsonx.data", "watsonx.ai (SaaS)"]},
    {"title": "Data Intelligence", "color": COLORS["HDR_BLUE"], "products": ["CP4D", "Data Product Hub", "Decision Optimization", "Knowledge Catalog", "Manta Data Lineage", "Optim & Master Data Mgmt", "SPSS Stats"]},
    {"title": "Data Integration", "color": COLORS["HDR_BLUE"], "products": ["CP4D", "Data Fabric", "Data Integration", "DataStage", "Databand", "Replication", "StreamSets"]},
    {"title": "Asset Lifecycle Management", "color": COLORS["HDR_PURPLE"], "products": ["EI", "Envizi", "HashiCorp Terraform", "Maximo", "Sterling Order & Inventory Mgmt", "Supply Chain", "TRIRIGA"]}
]

#    Row 3: App Dev & Int
CENTER_ROW_3 = [
    {"title": "Application Development", "color": COLORS["HDR_PURPLE"], "products": ["App Run", "CP4Apps", "CP4Systems", "DevOps", "ELM", "Project Harmony", "Runtimes", "Spectrum LSF", "UnifyBlue", "WAS", "WCA Java", "Web Hybrid ED"]},
    {"title": "Application Integration", "color": COLORS["HDR_PURPLE"], "products": ["API Connect", "APP Connect", "Aspera", "CP4I", "Connect:Direct", "DataPower", "DataPower Operational Dashboard", "Event Automation", "FTM", "MQ", "Sterling B2B Integrator", "WebMethods"]}
]

# 3. RIGHT COLUMN
RIGHT_COL = [
    {"title": "IT Automation & Finops", "color": COLORS["HDR_PURPLE"], "products": ["Ansible", "Apptio", "Cloud Pak for AIOps", "Cloudability", "Concert", "Flexera One", "HashiCorp Terraform", "Instana", "Kubecost", "Operations Insights", "Targetprocess", "Turbonomic", "Workload Automation"]},
    {"title": "Network Mgmt", "color": COLORS["HDR_PURPLE"], "products": ["CP4NA", "Cloud Network Security", "Content Delivery Network", "Edge Application Manager", "HashiCorp Nomad", "Hybrid Cloud Mesh", "NS1 Connect", "SevOne"]}
]

# 4. INFRASTRUCTURE (BOTTOM)
INFRA_COLS = [
    {"title": "Enterprise Storage", "color": COLORS["HDR_GREY"], "products": ["DS8000 Series", "SAN Directors", "Tape (Hydra & Jaguar)/VTS"]},
    {"title": "Data Resilience Storage", "color": COLORS["HDR_GREY"], "products": ["Scale", "Scale System", "Ceph", "CoS", "Defender/Protect", "Flash", "Fusion", "Fusion HCI", "Fusion HCI (on-prem)", "Hyperscaler (+Physical Tape)", "SVC", "Ceph System", "Storage Insight/Control", "Storage Virtualize", "Tape"]},
    {"title": "Power", "color": COLORS["HDR_GREY"], "products": ["AIX", "IBM i", "Linux", "Oracle", "Red Hat OpenShift", "SAP"]},
    {"title": "Z System", "color": COLORS["HDR_GREY"], "products": ["AI on Z", "IBM LinuxOne", "IBM zOS", "Z Monitoring Suite", "Z Security", "Z Software"]},
    {"title": "Cloud", "color": COLORS["HDR_GREY"], "products": ["Cloud Financial Server", "Cloud Satellite", "Power Virtual Server", "Red Hat OpenShift", "SAP", "VMware"]}
]

# --- HELPER FUNCTIONS ---

def create_box(slide, x, y, w, h, text, bg_color, font_color=COLORS["TEXT_STD"], 
               bold=False, font_size=9, outline_color=None, align=PP_ALIGN.CENTER):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    
    if outline_color:
        shape.line.color.rgb = outline_color
        shape.line.width = Pt(1.5)
    else:
        shape.line.fill.background() # No border if not specified

    tf = shape.text_frame
    tf.margin_top = Pt(1)
    tf.margin_bottom = Pt(1)
    tf.margin_left = Pt(1)
    tf.margin_right = Pt(1)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.name = FONT_NAME
    p.font.bold = bold
    p.font.color.rgb = font_color
    p.alignment = align
    return shape

def build_product_grid(slide, x, y, w, h, products, cols=1, border_color=COLORS["HDR_BLUE"]):
    """ Builds a grid of product boxes within a given area """
    gap = Inches(0.05)
    rows = (len(products) + cols - 1) // cols
    
    # Calculate box dimensions
    box_w = (w - (gap * (cols - 1))) / cols
    box_h = (h - (gap * (rows - 1))) / rows
    
    # Clamp max height to avoid huge boxes in sparse columns
    if box_h > Inches(0.4): box_h = Inches(0.4)

    for i, prod in enumerate(products):
        r = i // cols
        c = i % cols
        
        bx = x + (c * (box_w + gap))
        by = y + (r * (box_h + gap))
        
        # Special handling for "Sub-headers" inside Security column
        is_subheader = prod in ["Data Security", "Identity & Access Mgmt"]
        bg = COLORS["HDR_BLUE"] if is_subheader else COLORS["BOX_FILL"]
        fc = COLORS["TEXT_WHITE"] if is_subheader else COLORS["TEXT_STD"]
        
        create_box(slide, bx, by, box_w, box_h, prod, bg, font_color=fc, 
                   outline_color=border_color, font_size=8)

def build_pillar(slide, x, y, w, h, data, cols=1):
    """ Creates a header + product grid """
    header_h = Inches(0.35)
    
    # Header
    create_box(slide, x, y, w, header_h, data["title"], data["color"], 
               font_color=COLORS["TEXT_WHITE"], bold=True, font_size=10)
    
    # Body
    body_y = y + header_h + Inches(0.05)
    body_h = h - header_h - Inches(0.05)
    build_product_grid(slide, x, body_y, w, body_h, data["products"], cols, data["color"])

# --- MAIN LAYOUT BUILDER ---

prs = Presentation()
prs.slide_width = SLIDE_WIDTH
prs.slide_height = SLIDE_HEIGHT
slide = prs.slides.add_slide(prs.slide_layouts[6])

# 1. HEADER ROW
create_box(slide, Inches(0.2), Inches(0.2), Inches(3), Inches(0.5), "IBM Technology", COLORS["BG_HEADER"], font_size=24, bold=True, align=PP_ALIGN.LEFT)
# Dropdown placeholder
create_box(slide, Inches(3.5), Inches(0.3), Inches(2.5), Inches(0.3), "Default Placemat ▼", COLORS["BG_HEADER"], outline_color=COLORS["BORDER_GREY"], align=PP_ALIGN.LEFT)
# Button
create_box(slide, Inches(8.5), Inches(0.3), Inches(1.8), Inches(0.35), "Upload EPM Data", COLORS["BTN_BLUE"], font_color=COLORS["TEXT_WHITE"], bold=True)

# Legends
leg_w = Inches(1.5)
leg_x = SLIDE_WIDTH - (leg_w * 4) - Inches(0.5)
for txt, col in [("Entitled", "STS_ENTITLED"), ("Opportunity", "STS_OPP"), ("Explore", "STS_EXPLORE"), ("No Interest/At Risk", "STS_RISK")]:
    create_box(slide, leg_x, Inches(0.3), leg_w, Inches(0.35), txt, COLORS[col], font_size=10)
    leg_x += leg_w + Inches(0.1)

# 2. MAIN GRID SETUP
MARGIN_X = Inches(0.2)
TOP_Y = Inches(1.0)
LEFT_COL_W = Inches(2.0)
RIGHT_COL_W = Inches(2.0)
CENTER_W = SLIDE_WIDTH - LEFT_COL_W - RIGHT_COL_W - (MARGIN_X * 4)
GAP = Inches(0.1)

# X Positions
X_LEFT = MARGIN_X
X_CENTER = X_LEFT + LEFT_COL_W + GAP
X_RIGHT = X_CENTER + CENTER_W + GAP

# 3. LEFT COLUMN (Security)
# Security needs to be tall. Let's calculate total height based on Infrastructure position.
INFRA_H = Inches(2.5)
MAIN_H = SLIDE_HEIGHT - TOP_Y - INFRA_H - Inches(1.0) 

# Explicit split for Security sub-sections to match image visual
# Just rendering one big block for now as the data list combines them
build_pillar(slide, X_LEFT, TOP_Y, LEFT_COL_W, MAIN_H, LEFT_COL[0], cols=1)

# 4. RIGHT COLUMN (Auto & Net)
# Split Right column height 60/40
AUTO_H = MAIN_H * 0.6
NET_H = MAIN_H - AUTO_H - GAP
build_pillar(slide, X_RIGHT, TOP_Y, RIGHT_COL_W, AUTO_H, RIGHT_COL[0], cols=1)
build_pillar(slide, X_RIGHT, TOP_Y + AUTO_H + GAP, RIGHT_COL_W, NET_H, RIGHT_COL[1], cols=1)

# 5. CENTER STACK
# A. Client Apps (Top)
C_APP_H = Inches(1.5)
build_pillar(slide, X_CENTER, TOP_Y, CENTER_W, C_APP_H, CENTER_ROW_1, cols=8)

# B. The 6 Pillars (Middle)
PILLAR_Y = TOP_Y + C_APP_H + GAP
# Calculate remaining height for the 6 pillars and the bottom App row
REM_H = MAIN_H - C_APP_H - GAP
APP_ROW_H = Inches(2.2)
PILLAR_H = REM_H - APP_ROW_H - GAP

PILLAR_W = (CENTER_W - (GAP * 5)) / 6
for i, data in enumerate(CENTER_ROW_2):
    px = X_CENTER + (i * (PILLAR_W + GAP))
    build_pillar(slide, px, PILLAR_Y, PILLAR_W, PILLAR_H, data, cols=1)

# C. App Dev & Int (Bottom of Center)
APP_Y = PILLAR_Y + PILLAR_H + GAP
APP_HALF_W = (CENTER_W - GAP) / 2
build_pillar(slide, X_CENTER, APP_Y, APP_HALF_W, APP_ROW_H, CENTER_ROW_3[0], cols=4)
build_pillar(slide, X_CENTER + APP_HALF_W + GAP, APP_Y, APP_HALF_W, APP_ROW_H, CENTER_ROW_3[1], cols=4)

# 6. RED HAT OPENSHIFT BANNER
RH_Y = TOP_Y + MAIN_H + Inches(0.1)
RH_H = Inches(0.5)
create_box(slide, MARGIN_X, RH_Y, SLIDE_WIDTH - (MARGIN_X*2), RH_H, "Red Hat OpenShift", COLORS["BOX_FILL"], font_color=COLORS["HDR_RED"], outline_color=COLORS["HDR_RED"], bold=True, font_size=12)

# 7. INFRASTRUCTURE (Footer)
INFRA_Y = RH_Y + RH_H + Inches(0.1)
# 5 Columns
INFRA_COLS_W = (SLIDE_WIDTH - (MARGIN_X*2) - (GAP * 4)) / 5
for i, data in enumerate(INFRA_COLS):
    ix = MARGIN_X + (i * (INFRA_COLS_W + GAP))
    # Cols: Data Resilience needs 3, others 2
    c = 3 if "Resilience" in data["title"] else 2
    build_pillar(slide, ix, INFRA_Y, INFRA_COLS_W, INFRA_H, data, cols=c)

# 8. FOOTER LINKS
FOOT_Y = INFRA_Y + INFRA_H + Inches(0.1)
FOOT_W = (SLIDE_WIDTH - (MARGIN_X*2) - GAP) / 2
create_box(slide, MARGIN_X, FOOT_Y, FOOT_W, Inches(0.5), "IBM Technology Lifecycle Services (TLS)", RGBColor(240,240,240), outline_color=COLORS["BORDER_GREY"], bold=True)
create_box(slide, MARGIN_X + FOOT_W + GAP, FOOT_Y, FOOT_W, Inches(0.5), "IBM Expert Labs (EL)", RGBColor(240,240,240), outline_color=COLORS["BORDER_GREY"], bold=True)

# SAVE
prs.save("IBM_Product_Placemat.pptx")
print("Slide generated successfully.")
