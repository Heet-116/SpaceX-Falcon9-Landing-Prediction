import os
import sqlite3
import pandas as pd
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# 1. Database connection and SQL execution
df_sql = pd.read_csv('data/Spacex.csv')
conn = sqlite3.connect(':memory:')
df_sql.to_sql('SPACEXTBL', conn, index=False, if_exists='replace')

sql_queries = {
    24: ("SELECT DISTINCT Launch_Site FROM SPACEXTBL",
         "The dataset contains 4 distinct launch sites. CCAFS LC-40 and KSC LC-39A serve East Coast launches into equatorial and high-inclination orbits, while VAFB SLC-4E serves West Coast launches into polar orbits."),
    25: ("SELECT Date, Booster_Version, Launch_Site, PAYLOAD_MASS__KG_, Landing_Outcome FROM SPACEXTBL WHERE Launch_Site LIKE 'CCA%' LIMIT 5",
         "The first 5 orbital flights were conducted from Cape Canaveral Air Force Station (CCAFS LC-40) carrying experimental Dragon qualification payloads with ocean splashdown landing attempts."),
    26: ("SELECT SUM(PAYLOAD_MASS__KG_) AS Total_Payload_NASA_CRS_KG FROM SPACEXTBL WHERE Customer = 'NASA (CRS)'",
         "SpaceX delivered a cumulative total of 45,596 kg of cargo to the International Space Station under NASA Commercial Resupply Services (CRS) contracts across 20+ missions."),
    27: ("SELECT ROUND(AVG(PAYLOAD_MASS__KG_), 2) AS Avg_Payload_F9_v1_1_KG FROM SPACEXTBL WHERE Booster_Version LIKE 'F9 v1.1%'",
         "The Falcon 9 v1.1 variant carried an average payload of 2,928.4 kg, serving early satellite constellations and ISS resupply prior to the introduction of Full Thrust and Block 5."),
    28: ("SELECT MIN(Date) AS First_Successful_Ground_Landing, Booster_Version, Launch_Site, Landing_Outcome FROM SPACEXTBL WHERE Landing_Outcome = 'Success (ground pad)'",
         "The historic first successful orbital first-stage recovery occurred on December 22, 2015, by booster B1019 at Cape Canaveral Landing Zone 1 (LZ-1), validating rocket reusability."),
    29: ("SELECT Booster_Version, PAYLOAD_MASS__KG_, Launch_Site FROM SPACEXTBL WHERE Landing_Outcome = 'Success (drone ship)' AND PAYLOAD_MASS__KG_ BETWEEN 4000 AND 6000",
         "Four boosters (B1022, B1026, B1021.2, B1031.2) successfully completed Autonomous Spaceport Drone Ship (ASDS) recoveries during demanding GTO missions where return-to-launch-site was fuel-prohibitive."),
    30: ("SELECT Mission_Outcome, COUNT(*) AS Outcome_Count FROM SPACEXTBL GROUP BY Mission_Outcome",
         "Out of 101 recorded missions, 98 were completely successful primary missions (97.0% mission success rate), demonstrating industrial reliability across civil, defense, and commercial launches."),
    31: ("SELECT Booster_Version, Launch_Site, PAYLOAD_MASS__KG_ FROM SPACEXTBL WHERE PAYLOAD_MASS__KG_ = (SELECT MAX(PAYLOAD_MASS__KG_) FROM SPACEXTBL) LIMIT 5",
         "The maximum payload mass recorded is 15,600 kg, carried by Falcon 9 Block 5 boosters (e.g. B1048, B1049, B1051) launching dense batches of 60 Starlink satellites to Low Earth Orbit."),
    32: ("SELECT Date, Booster_Version, Launch_Site, Landing_Outcome FROM SPACEXTBL WHERE Landing_Outcome = 'Failure (drone ship)' AND Date LIKE '2015%'",
         "In 2015, two drone ship landing attempts failed (B1012 on Jan 10 and B1015 on Apr 14) due to hydraulic fluid exhaustion and stick-slip thruster stiction, providing vital telemetry for subsequent fixes."),
    33: ("SELECT Landing_Outcome, COUNT(*) AS Total_Count FROM SPACEXTBL WHERE Date BETWEEN '2010-06-04' AND '2017-03-20' GROUP BY Landing_Outcome ORDER BY Total_Count DESC",
         "Between 2010 and 2017, early operations included 10 'No attempt' missions, followed by 5 drone ship successes, 5 drone ship failures, and 3 ground pad successes as landing technology matured.")
}

sql_results_data = {}
for slide_idx, (q, expl) in sql_queries.items():
    df_res = pd.read_sql_query(q, conn)
    sql_results_data[slide_idx] = (q, df_res, expl)

print("SQL queries successfully executed.")

# 2. Presentation Setup
prs = pptx.Presentation('ds-capstone-template-coursera.pptx')
SW = prs.slide_width
SH = prs.slide_height

# Styling Constants
COLOR_TITLE = RGBColor(24, 43, 73)      # Dark Navy
COLOR_TEXT = RGBColor(44, 62, 80)       # Slate Dark Gray
COLOR_MUTED = RGBColor(127, 140, 141)   # Muted Gray
COLOR_HIGHLIGHT = RGBColor(41, 128, 185)# SpaceX Blue
COLOR_ACCENT = RGBColor(39, 174, 96)    # Forest Green
COLOR_BG_BOX = RGBColor(245, 247, 250)  # Light Card Background
COLOR_QUERY_BG = RGBColor(15, 23, 42)   # Dark Terminal Slate
COLOR_WHITE = RGBColor(255, 255, 255)
FONT_NAME = 'Calibri'
WATERMARK_TEXT = "Heet Patel | IBM Applied Data Science Capstone | https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction"

def add_watermark(slide):
    tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.15), Inches(11.5), Inches(0.28))
    tf = tx_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = WATERMARK_TEXT
    p.font.name = FONT_NAME
    p.font.size = Pt(8.5)
    p.font.color.rgb = COLOR_MUTED

def set_title(slide, title_text):
    if slide.shapes.title:
        slide.shapes.title.text = title_text
        for p in slide.shapes.title.text_frame.paragraphs:
            p.font.name = FONT_NAME
            p.font.size = Pt(24)
            p.font.bold = True
            p.font.color.rgb = COLOR_TITLE

def populate_text_bullets(shape, bullet_groups, font_size=13):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    
    first = True
    for header, bullets in bullet_groups:
        if header:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.text = header
            p.font.name = FONT_NAME
            p.font.size = Pt(font_size + 1)
            p.font.bold = True
            p.font.color.rgb = COLOR_HIGHLIGHT
            p.space_after = Pt(3)
            p.space_before = Pt(6) if not first else Pt(0)
            
        for b in bullets:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.text = f"• {b}"
            p.font.name = FONT_NAME
            p.font.size = Pt(font_size)
            p.font.color.rgb = COLOR_TEXT
            p.space_after = Pt(4)

print("\n--- Step 3: Populating All Slides ---")

# SLIDE 1: Title Slide
s1 = prs.slides[0]
for sh in s1.shapes:
    if sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(2.2)
        sh.width = Inches(11.5)
        sh.height = Inches(4.5)
        tf = sh.text_frame
        tf.clear()
        
        p1 = tf.paragraphs[0]
        p1.text = "SpaceX Falcon 9 First Stage Landing Prediction"
        p1.font.name = FONT_NAME
        p1.font.size = Pt(32)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_TITLE
        p1.space_after = Pt(10)
        
        p2 = tf.add_paragraph()
        p2.text = "End-to-End Data Science & Machine Learning Capstone Report"
        p2.font.name = FONT_NAME
        p2.font.size = Pt(20)
        p2.font.color.rgb = COLOR_HIGHLIGHT
        p2.space_after = Pt(24)
        
        p3 = tf.add_paragraph()
        p3.text = "Author: Heet Patel | IBM Data Science Professional Certificate"
        p3.font.name = FONT_NAME
        p3.font.size = Pt(14)
        p3.font.bold = True
        p3.font.color.rgb = COLOR_TEXT
        p3.space_after = Pt(4)
        
        p4 = tf.add_paragraph()
        p4.text = "Date: October 2026 | Submission Format: PDF & Python/Jupyter Deliverables"
        p4.font.name = FONT_NAME
        p4.font.size = Pt(13)
        p4.font.color.rgb = COLOR_MUTED
        p4.space_after = Pt(8)
        
        p5 = tf.add_paragraph()
        p5.text = "GitHub Repository: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction"
        p5.font.name = FONT_NAME
        p5.font.size = Pt(13)
        p5.font.bold = True
        p5.font.color.rgb = COLOR_ACCENT

# SLIDE 2: Outline
s2 = prs.slides[1]
set_title(s2, "Table of Contents & Project Roadmap")
add_watermark(s2)
for sh in s2.shapes:
    if sh != s2.shapes.title and sh.has_text_frame:
        populate_text_bullets(sh, [
            ("Project Overview & Foundations", [
                "1. Executive Summary: Core findings, pipeline summary, and commercial takeaways",
                "2. Introduction & Background: Reusable rocketry economics and research questions"
            ]),
            ("Methodologies & Technical Architecture", [
                "3. Data Collection & Wrangling: SpaceX REST API, Wikipedia scraping, and data cleaning",
                "4. Exploratory Data Analysis: Data visualization (Matplotlib/Seaborn) and SQL querying (SQLite)",
                "5. Interactive Analytics: Folium geospatial mapping and Plotly Dash dynamic web dashboard",
                "6. Predictive Modeling: Classification algorithms, hyperparameter tuning, and model evaluation"
            ]),
            ("Results, Conclusions & Appendix", [
                "7. Complete Results: Visualizations, SQL outputs, interactive maps, and ML confusion matrix",
                "8. Conclusions & Business Impact: Competitive launch bidding and commercial strategy",
                "9. Appendix: Data dictionary, technology stack, and verified GitHub repository code links"
            ])
        ], font_size=12)

# SLIDE 3: Executive Summary
s3 = prs.slides[2]
set_title(s3, "Executive Summary: Methodology & Key Results")
add_watermark(s3)
for sh in s3.shapes:
    if sh != s3.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(11.5)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Summary of Methodologies", [
                "Data Collection & Processing: Ingested 90 Falcon 9 launches via SpaceX REST API v4 and historical Wikipedia web scraping using BeautifulSoup. Cleaned null payload records with mean imputation (6,105 kg) and engineered a binary 'Class' landing target (1 = Success, 0 = Failure).",
                "Exploratory Data Analysis: Analyzed relationships between payload, orbit, flight number, and launch site; executed 10 structured SQL queries in SQLite to isolate operational milestones, payloads, and booster variants.",
                "Interactive Analytics: Built Folium maps with marker clusters and Haversine proximity calculations to coastal/transport infrastructure; engineered a responsive Plotly Dash dashboard with payload range sliders and site dropdowns.",
                "Machine Learning: Standardized 83 features via StandardScaler; trained 4 classification models (Logistic Regression, SVM, Decision Tree, KNN) optimized using 10-fold cross-validation GridSearchCV."
            ]),
            ("Summary of Key Results & Commercial Takeaways", [
                "Overall Landing Success Rate: 66.7% across all 90 launches; success escalated from 0% (2010–2012) to >80% (2018–2020) driven by Falcon 9 Block 5 reusability advancements.",
                "Top Launch Site: KSC LC-39A achieved the highest success rate (76.9%) and contributed 41.7% of all successful first-stage recoveries.",
                "Predictive Accuracy: All 4 machine learning models achieved 83.33% Test Accuracy. The models delivered 100% Recall for successful landings (0 False Negatives), preventing premature disposal of recoverable boosters.",
                "Business Impact: Enables commercial satellite operators and competitor launch providers to predict SpaceX booster recovery with high confidence, informing competitive pricing against SpaceX's $62M launch rate."
            ])
        ], font_size=11)

# SLIDE 4: Introduction
s4 = prs.slides[3]
set_title(s4, "Introduction: Background & Core Objectives")
add_watermark(s4)
for sh in s4.shapes:
    if sh != s4.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(11.5)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Project Background & Commercial Context", [
                "SpaceX Revolution: SpaceX has drastically reduced orbital launch costs by mastering first-stage booster recovery and reuse.",
                "Economic Asymmetry: Competitor launch providers (e.g., ULA, Arianespace, Roscosmos) operate expendable rockets costing upwards of $165M+ per launch, whereas SpaceX advertises Falcon 9 launches at approximately $62M.",
                "Value of Recovery: The first-stage booster represents ~70% of total rocket manufacturing cost. Reusing boosters multiple times amortizes production capital and yields unprecedented commercial margins."
            ]),
            ("Core Research Questions to Answer", [
                "1. Determinants of Recovery: How do payload mass, target orbital profile, launch site geography, and booster flight cadence influence landing success?",
                "2. Geospatial Patterns: Why are launch complexes positioned along specific coastlines, and how do distances to transport and cities affect operational safety?",
                "3. Machine Learning Viability: Can supervised classification models accurately forecast landing outcomes for future launches prior to liftoff?",
                "4. Strategic Bidding Utility: How can competitor launch providers utilize these predictive models to formulate competitive bids against SpaceX contracts?"
            ])
        ], font_size=11.5)

# SLIDE 6: Methodology Overview
s6 = prs.slides[5]
set_title(s6, "End-to-End Methodology Architecture")
add_watermark(s6)
for sh in s6.shapes:
    if sh != s6.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(11.5)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Phase 1: Data Collection & Ingestion", [
                "Extracted 90+ rocket launch records via SpaceX REST API v4 (JSON) capturing cores, flight numbers, payload specs, and landing pads.",
                "Scraped historical Falcon 9 launch logs from Wikipedia using BeautifulSoup to extract customers, dates, and mission outcomes."
            ]),
            ("Phase 2: Data Wrangling & Feature Engineering", [
                "Identified missing payload values and imputed with mean (6,104.96 kg); mapped raw landing outcomes into binary classification target 'Class'.",
                "Applied one-hot encoding on categorical features (Orbit, LaunchSite, LandingPad, Serial), expanding the feature space to 83 numerical columns."
            ]),
            ("Phase 3: Exploratory Data Analysis & Relational Querying", [
                "Executed visualization analysis across payload, orbit, and temporal trends using Matplotlib and Seaborn.",
                "Loaded datasets into SQLite database 'SPACEXTBL' and conducted 10 technical SQL queries for aggregated insights and historical milestones."
            ]),
            ("Phase 4: Interactive Visual Analytics", [
                "Constructed Folium interactive maps with marker clusters and Haversine distance proximity measurements to coastline, transport, and cities.",
                "Engineered a dynamic Plotly Dash web dashboard featuring reactive dropdowns and payload range sliders for live slice-and-dice analytics."
            ]),
            ("Phase 5: Predictive Modeling & Hyperparameter Tuning", [
                "Standardized features via StandardScaler; partitioned data into 80% Train and 20% Test sets (random_state=2).",
                "Tuned Logistic Regression, SVM, Decision Tree, and KNN models using GridSearchCV with 10-fold cross validation; evaluated test accuracy and confusion matrix."
            ])
        ], font_size=10.5)

# SLIDE 7: Data Collection Overview
s7 = prs.slides[6]
set_title(s7, "Data Collection Methodology: REST API & Web Scraping")
add_watermark(s7)
for sh in s7.shapes:
    if sh != s7.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(11.5)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Data Source 1: SpaceX REST API (JSON)", [
                "Endpoint: https://api.spacexdata.com/v4/launches/past",
                "Extraction Method: Used Python 'requests' library to pull comprehensive raw launch data in JSON format.",
                "Data Structure: Utilized 'pd.json_normalize' to unnest JSON payloads, extracting rocket IDs, cores, payloads, flight numbers, and launchpads.",
                "Filtering: Filtered dataset strictly for Falcon 9 launches (excluding Falcon 1 test flights and Falcon Heavy 3-core configurations), yielding 90 launches."
            ]),
            ("Data Source 2: Wikipedia Web Scraping (HTML Tables)", [
                "Target Page: 'List of Falcon 9 and Falcon Heavy launches' Wikipedia historical archive.",
                "Extraction Method: Applied BeautifulSoup to parse HTML document trees and navigate 'wikitable' table structures.",
                "Extracted Attributes: Extracted Flight Number, Launch Date, Booster Version, Launch Site, Payload Name, Payload Mass, Orbit Type, Customer, and Outcome.",
                "Data Sanitization: Stripped superscript citations (e.g. '[12]'), handled irregular dates, and standardized numeric payload masses."
            ]),
            ("Data Synthesis & Quality Assurance", [
                "Cross-referenced API records with Wikipedia launch logs to ensure zero missing core serials and full alignment on mission outcomes.",
                "GitHub Notebook References: [1_SpaceX_Data_Collection_API.ipynb] & [2_SpaceX_Data_Collection_Web_Scraping.ipynb]"
            ])
        ], font_size=11)

# SLIDE 8: Data Collection – SpaceX API (2-column: Text Left, Flowchart Right)
s8 = prs.slides[7]
set_title(s8, "Data Collection: SpaceX REST API Process")
add_watermark(s8)
for sh in s8.shapes:
    if sh != s8.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(5.2)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("API Call Process & Telemetry Extraction", [
                "Endpoint: api.spacexdata.com/v4/launches/past",
                "Payload Details: Retrieved payload masses, destination orbits, and customer IDs from child endpoints.",
                "Core Telemetry: Extracted booster serials, flight counts, gridfins status, landing leg deployment, landing pad IDs, and landing outcomes.",
                "Dataframe Construction: Consolidated unnested records into a structured Pandas DataFrame of 90 Falcon 9 missions with 17 features."
            ]),
            ("GitHub Notebook Reference", [
                "Completed Notebook: notebooks/1_SpaceX_Data_Collection_API.ipynb",
                "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/1_SpaceX_Data_Collection_API.ipynb"
            ])
        ], font_size=11)
s8.shapes.add_picture('images/flowchart_api.png', Inches(6.2), Inches(1.6), Inches(5.2), Inches(4.8))

# SLIDE 9: Data Collection – Scraping (2-column: Text Left, Flowchart Right)
s9 = prs.slides[8]
set_title(s9, "Data Collection: Wikipedia Web Scraping Process")
add_watermark(s9)
for sh in s9.shapes:
    if sh != s9.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(5.2)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("HTML Scraping Pipeline with BeautifulSoup", [
                "Target URL: Wikipedia List of Falcon 9 launches",
                "HTML Parsing: Inspected DOM hierarchy; targeted table elements with class 'wikitable'.",
                "Table Traversal: Extracted column headers from <th> tags; iterated through <tr> rows and <td> cells.",
                "Cleaning Rules: Regex used to strip citation brackets ('[1]'), remove unicode whitespace, and cast payload masses to float.",
                "Output: Built a secondary reference DataFrame matching historical launch records."
            ]),
            ("GitHub Notebook Reference", [
                "Completed Notebook: notebooks/2_SpaceX_Data_Collection_Web_Scraping.ipynb",
                "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/2_SpaceX_Data_Collection_Web_Scraping.ipynb"
            ])
        ], font_size=11)
s9.shapes.add_picture('images/flowchart_scraping.png', Inches(6.2), Inches(1.6), Inches(5.2), Inches(4.8))

# SLIDE 10: Data Wrangling (2-column: Text Left, Flowchart Right)
s10 = prs.slides[9]
set_title(s10, "Data Wrangling & Feature Engineering")
add_watermark(s10)
for sh in s10.shapes:
    if sh != s10.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(5.2)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Data Cleaning & Target Encoding", [
                "Handling Nulls: Identified 5 missing values in 'PayloadMass'; imputed with mean (6,104.96 kg) to maintain sample size.",
                "Outcome Categorization: Categorized landing outcomes: 'True RTLS', 'True ASDS', 'True Ocean' mapped to Class = 1 (Success); 'False RTLS/ASDS/Ocean' and 'None' mapped to Class = 0 (Failure).",
                "Target Balance: 60 Successful landings (66.7%), 30 Failed/Unattempted landings (33.3%).",
                "One-Hot Encoding: Converted categorical variables ('Orbit', 'LaunchSite', 'LandingPad', 'Serial') into 83 dummy columns using pd.get_dummies."
            ]),
            ("GitHub Notebook Reference", [
                "Completed Notebook: notebooks/3_SpaceX_Data_Wrangling.ipynb",
                "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/3_SpaceX_Data_Wrangling.ipynb"
            ])
        ], font_size=10.5)
s10.shapes.add_picture('images/flowchart_wrangling.png', Inches(6.2), Inches(1.6), Inches(5.2), Inches(4.8))

# SLIDE 11: EDA with Data Visualization Methodology
s11 = prs.slides[10]
set_title(s11, "EDA with Data Visualization Methodology")
add_watermark(s11)
for sh in s11.shapes:
    if sh != s11.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(11.5)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Exploratory Data Visualization Pipeline", [
                "Objective: Uncover hidden patterns, correlations, and success drivers across physical, geographic, and temporal flight variables.",
                "Tooling: Python Data Science Stack (Pandas, NumPy, Matplotlib, Seaborn).",
                "Key Visualizations Generated:",
                "  1. Flight Number vs. Launch Site: Scatter plot tracking site migration and operational maturity over time.",
                "  2. Payload Mass vs. Launch Site: Scatter plot assessing site-specific payload lift capacity and recovery success.",
                "  3. Success Rate vs. Orbit Type: Categorical bar chart quantifying risk levels across 11 distinct orbital profiles.",
                "  4. Flight Number vs. Orbit Type: Scatter plot mapping cadence growth from early LEO/GTO flights to high-volume VLEO constellations.",
                "  5. Payload Mass vs. Orbit Type: Scatter plot analyzing energy-payload boundaries and booster fuel recovery feasibility.",
                "  6. Yearly Success Trend: Temporal line chart tracing SpaceX's engineering learning curve from 2010 to 2020."
            ]),
            ("GitHub Notebook Reference", [
                "Completed Notebook: notebooks/4_SpaceX_EDA_Data_Visualization.ipynb",
                "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/4_SpaceX_EDA_Data_Visualization.ipynb"
            ])
        ], font_size=11)

# SLIDE 12: EDA with SQL Methodology
s12 = prs.slides[11]
set_title(s12, "EDA with SQL Querying Methodology")
add_watermark(s12)
for sh in s12.shapes:
    if sh != s12.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(11.5)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Relational SQL Analysis Environment", [
                "Database Engine: Loaded SpaceX historical launch dataset (101 records) into an SQLite database table named 'SPACEXTBL'.",
                "Querying Objectives: Conduct deterministic data profiling, historical milestone extraction, and customer payload auditing.",
                "Core SQL Capabilities Implemented:",
                "  • Filtering & Pattern Matching: Applied LIKE 'CCA%' to isolate Cape Canaveral launches, and LIKE 'F9 v1.1%' to filter booster versions.",
                "  • Mathematical Aggregations: Computed SUM(PAYLOAD_MASS__KG_) for NASA CRS contracts and AVG(PAYLOAD_MASS__KG_) across booster variants.",
                "  • Chronological Extraction: Used MIN(Date) to pinpoint the exact date of the world's first successful orbital rocket landing.",
                "  • Subqueries & Extrema: Formatted nested subqueries (WHERE PAYLOAD_MASS__KG_ = (SELECT MAX(...))) to identify heaviest liftoffs.",
                "  • Frequency Ranking: Grouped outcomes by type, filtered date ranges using BETWEEN, and ordered by descending count."
            ]),
            ("GitHub Notebook Reference", [
                "Completed Notebook: notebooks/5_SpaceX_EDA_SQL.ipynb",
                "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/5_SpaceX_EDA_SQL.ipynb"
            ])
        ], font_size=11)

# SLIDE 13: Build an Interactive Map with Folium Methodology
s13 = prs.slides[12]
set_title(s13, "Interactive Visual Analytics with Folium Methodology")
add_watermark(s13)
for sh in s13.shapes:
    if sh != s13.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(11.5)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Folium Geospatial Mapping Architecture", [
                "Geographic Engine: Utilized Python 'folium' library to render dynamic Leaflet.js interactive maps centered on US launch complexes.",
                "Marker Architecture: Generated 'folium.Marker' and 'folium.Circle' objects for CCAFS LC-40, CCAFS SLC-40, KSC LC-39A, and VAFB SLC-4E.",
                "Marker Clustering Plugin: Implemented 'folium.plugins.MarkerCluster' to aggregate dense launch markers, preventing visual clutter.",
                "Color-Coded Telemetry: Created custom icon markers color-coded by landing outcome (Green = Landing Success, Red = Landing Failure).",
                "Proximity & Distance Calculations: Implemented the Haversine spherical trigonometric formula to calculate exact straight-line distances from launch pads to coastlines, railways, highways, and nearest cities, overlaying dashed 'folium.PolyLine' visualizers."
            ]),
            ("GitHub Notebook Reference", [
                "Completed Notebook: notebooks/6_SpaceX_Interactive_Visual_Analytics_Folium.ipynb",
                "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/6_SpaceX_Interactive_Visual_Analytics_Folium.ipynb"
            ])
        ], font_size=11)

# SLIDE 14: Build a Dashboard with Plotly Dash Methodology
s14 = prs.slides[13]
set_title(s14, "Interactive Dashboard Architecture with Plotly Dash")
add_watermark(s14)
for sh in s14.shapes:
    if sh != s14.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(11.5)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Plotly Dash Web Application Architecture", [
                "Dashboard Framework: Built a multi-component reactive web application using Dash, dash_core_components (dcc), and dash_html_components (html).",
                "Interactive UI Controls:",
                "  • Dropdown Component (dcc.Dropdown): Allows user to select 'All Sites' or toggle individual launch sites (CCAFS LC-40, VAFB SLC-4E, KSC LC-39A, CCAFS SLC-40).",
                "  • Range Slider Component (dcc.RangeSlider): Enables fluid filtering across the entire payload spectrum (0 to 10,000 kg with 1,000 kg intervals).",
                "Reactive Callback Logic (@app.callback):",
                "  • Callback 1: Generates dynamic 'px.pie' charts displaying total success proportion across all sites, or success vs. failure ratio for a selected site.",
                "  • Callback 2: Dynamically renders 'px.scatter' plots of Payload Mass vs. Launch Outcome, color-coded by Booster Version Category."
            ]),
            ("GitHub Script Reference", [
                "Completed Application Script: spacex_dash_app.py",
                "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/spacex_dash_app.py"
            ])
        ], font_size=11)

# SLIDE 15: Predictive Analysis Methodology (2-column: Text Left, Flowchart Right)
s15 = prs.slides[14]
set_title(s15, "Predictive Analysis: Machine Learning Methodology")
add_watermark(s15)
for sh in s15.shapes:
    if sh != s15.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(5.2)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Machine Learning Development Pipeline", [
                "Feature Standardization: Standardized 83 numerical features using Scikit-Learn 'StandardScaler' (zero mean, unit variance).",
                "Data Partitioning: Split into 80% Training (72 samples) and 20% Testing (18 samples) with random_state=2.",
                "Hyperparameter Tuning: Tuned 4 classifiers via GridSearchCV with 10-fold cross validation (cv=10):",
                "  • Logistic Regression: L2 penalty, lbfgs solver, C in [0.01, 0.1, 1].",
                "  • Support Vector Machine: Linear, RBF, Poly, Sigmoid kernels; C & gamma grids.",
                "  • Decision Tree: Gini & entropy criteria; max_depth 2 to 18; min_samples_leaf.",
                "  • K-Nearest Neighbors: k from 1 to 10; Euclidean and Manhattan distance metrics.",
                "Evaluation: Measured test set accuracy, confusion matrices, precision, and recall."
            ]),
            ("GitHub Notebook Reference", [
                "Completed Notebook: notebooks/7_SpaceX_Machine_Learning_Prediction.ipynb",
                "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/7_SpaceX_Machine_Learning_Prediction.ipynb"
            ])
        ], font_size=10.5)
s15.shapes.add_picture('images/flowchart_ml.png', Inches(6.2), Inches(1.6), Inches(5.2), Inches(4.8))

# SLIDE 18: Flight Number vs. Launch Site
s18 = prs.slides[17]
set_title(s18, "EDA: Flight Number vs. Launch Site")
add_watermark(s18)
for sh in s18.shapes:
    if sh != s18.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Key Observations & Flight Cadence", [
                "Site Progression: Flights 1 through ~25 were launched almost exclusively from CCAFS LC-40, reflecting the initial experimental phase of Falcon 9.",
                "Landing Failure Clustering: Most early landing failures (red points) occurred at CCAFS LC-40 during early recovery development (2010–2015).",
                "Operational Maturity: Flights beyond #30 diversified to VAFB SLC-4E and KSC LC-39A, showing a dramatic concentration of green landing successes.",
                "Modern Reliability: For flight numbers >50, landing success rate exceeded 80% across all launch sites."
            ])
        ], font_size=11)
s18.shapes.add_picture('images/eda_flight_vs_launchsite.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 19: Payload vs. Launch Site
s19 = prs.slides[18]
set_title(s19, "EDA: Payload Mass vs. Launch Site")
add_watermark(s19)
for sh in s19.shapes:
    if sh != s19.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Key Observations & Payload Distribution", [
                "VAFB SLC-4E Profile: Specialized in lighter to medium payload missions (<10,000 kg) dedicated to polar and sun-synchronous orbits, achieving consistent landing success.",
                "KSC LC-39A Dominance: Handled both medium and heavy payloads (>10,000 kg), including commercial communications satellites and ISS Dragon resupply, with an outstanding recovery record.",
                "CCAFS SLC-40 Range: Dispersed across the full payload spectrum (0 to 15,600 kg); early low-mass flights had mixed results, while modern heavy missions achieved consistent recovery."
            ])
        ], font_size=11)
s19.shapes.add_picture('images/eda_payload_vs_launchsite.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 20: Success Rate vs. Orbit Type
s20 = prs.slides[19]
set_title(s20, "EDA: Landing Success Rate by Orbit Type")
add_watermark(s20)
for sh in s20.shapes:
    if sh != s20.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Key Observations by Orbital Trajectory", [
                "100% Success Orbits: ES-L1 (Earth-Sun Lagrange), GEO (Geosynchronous), HEO (Highly Elliptical), and SSO (Sun-Synchronous) all achieved a flawless 100% landing success rate.",
                "VLEO Constellations: Very Low Earth Orbit (Starlink) achieved ~85% success across high flight frequencies.",
                "LEO & Polar Orbits: LEO achieved ~71% and Polar Orbits (PO) achieved ~67% success rate.",
                "GTO Challenges: Geostationary Transfer Orbit achieved ~50% success due to extreme downrange velocity and severe thermal stress during atmospheric entry.",
                "SO Anomalies: Sub-Orbital test flight experienced 0% landing success."
            ])
        ], font_size=11)
s20.shapes.add_picture('images/eda_success_rate_vs_orbit.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 21: Flight Number vs. Orbit Type
s21 = prs.slides[20]
set_title(s21, "EDA: Flight Number vs. Orbit Type")
add_watermark(s21)
for sh in s21.shapes:
    if sh != s21.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Key Observations & Mission Evolution", [
                "Early Orbits: Early flight numbers (<30) primarily targeted LEO and GTO orbits with experimental landing profiles.",
                "Constellation Ramp-Up: Beyond Flight 50, mission frequency shifted dramatically toward VLEO (Starlink satellite mega-constellation deployment).",
                "Recovery Consistency: While early LEO and GTO missions experienced landing failures, later flights across all orbit types show nearly unbroken runs of green landing successes."
            ])
        ], font_size=11)
s21.shapes.add_picture('images/eda_flight_vs_orbit.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 22: Payload vs. Orbit Type
s22 = prs.slides[21]
set_title(s22, "EDA: Payload Mass vs. Orbit Type")
add_watermark(s22)
for sh in s22.shapes:
    if sh != s22.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Key Observations & Orbital Payload Limits", [
                "Heavy LEO/VLEO Missions: Maximum payload masses (>10,000 kg up to 15,600 kg) are concentrated in LEO and VLEO orbits, where landing success is consistently high.",
                "GTO Payload Constraint: GTO missions cluster in the 3,000 kg to 7,000 kg range. Heavy GTO launches leave less reserve fuel for the first-stage boostback and entry burns.",
                "Polar & Scientific Missions: Polar and SSO missions carried payloads under 4,000 kg, allowing comfortable fuel margins for booster recovery."
            ])
        ], font_size=11)
s22.shapes.add_picture('images/eda_payload_vs_orbit.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 23: Launch Success Yearly Trend
s23 = prs.slides[22]
set_title(s23, "EDA: Yearly Launch Success Rate Trend (2010–2020)")
add_watermark(s23)
for sh in s23.shapes:
    if sh != s23.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Key Observations & Technology Evolution", [
                "2010–2012 (0% Success): Purely expendable test launches; no booster recovery hardware existed.",
                "2013–2015 (~33% Success): Controlled ocean splashdowns and drone ship tests began; historic first ground pad recovery on Dec 22, 2015.",
                "2016–2017 (60–80% Success): Routine drone ship recoveries established on Autonomous Spaceport Drone Ships (ASDS).",
                "2018–2020 (>80–85% Success): Introduction of Falcon 9 Block 5 (strengthened heat shield, titanium grid fins, upgraded engines) established industrial recovery reliability."
            ])
        ], font_size=11)
s23.shapes.add_picture('images/eda_yearly_trend.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# HELPER: Function to format SQL Slides (24 - 33)
def build_sql_slide(slide, slide_num, title_text):
    set_title(slide, title_text)
    add_watermark(slide)
    
    query, df_res, explanation = sql_results_data[slide_num]
    
    # Remove existing text boxes except title and slide number
    for sh in list(slide.shapes):
        if sh != slide.shapes.title and sh.name.startswith(('Content', 'Text')):
            # Clear text frame
            sh.text_frame.clear()
            # Repurpose shape for query card
            sh.left = Inches(0.8)
            sh.top = Inches(1.3)
            sh.width = Inches(11.5)
            sh.height = Inches(0.9)
            tf = sh.text_frame
            tf.word_wrap = True
            p0 = tf.paragraphs[0]
            p0.text = "SQL Query Syntax:"
            p0.font.name = FONT_NAME
            p0.font.size = Pt(11)
            p0.font.bold = True
            p0.font.color.rgb = COLOR_HIGHLIGHT
            p1 = tf.add_paragraph()
            p1.text = query
            p1.font.name = 'Consolas'
            p1.font.size = Pt(10.5)
            p1.font.bold = True
            p1.font.color.rgb = COLOR_QUERY_BG
            break
            
    # Add Query Results Table
    num_rows = min(len(df_res) + 1, 9)
    num_cols = len(df_res.columns)
    
    tbl_left = Inches(0.8)
    tbl_top = Inches(2.35)
    tbl_width = Inches(11.5)
    tbl_height = Inches(0.35 * num_rows)
    
    table_shape = slide.shapes.add_table(num_rows, num_cols, tbl_left, tbl_top, tbl_width, tbl_height)
    table = table_shape.table
    
    # Format Headers
    for c_idx, col in enumerate(df_res.columns):
        cell = table.cell(0, c_idx)
        cell.text = str(col)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_TITLE
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_NAME
            p.font.size = Pt(10.5)
            p.font.bold = True
            p.font.color.rgb = COLOR_WHITE
            
    # Format Data Rows
    for r_idx in range(num_rows - 1):
        for c_idx in range(num_cols):
            cell = table.cell(r_idx + 1, c_idx)
            val = df_res.iloc[r_idx, c_idx]
            cell.text = str(val) if pd.notnull(val) else "NULL"
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(248, 250, 252) if r_idx % 2 == 0 else COLOR_WHITE
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                p.font.name = FONT_NAME
                p.font.size = Pt(10)
                p.font.color.rgb = COLOR_TEXT

    # Add Analytical Explanation Card below table
    card_top = tbl_top + tbl_height + Inches(0.2)
    card_height = Inches(6.8) - card_top
    if card_height > Inches(0.7):
        exp_box = slide.shapes.add_textbox(Inches(0.8), card_top, Inches(11.5), card_height)
        tf_e = exp_box.text_frame
        tf_e.word_wrap = True
        p_eh = tf_e.paragraphs[0]
        p_eh.text = "Key Analytical & Business Insight:"
        p_eh.font.name = FONT_NAME
        p_eh.font.size = Pt(11)
        p_eh.font.bold = True
        p_eh.font.color.rgb = COLOR_HIGHLIGHT
        p_eh.space_after = Pt(2)
        
        p_eb = tf_e.add_paragraph()
        p_eb.text = f"• {explanation}"
        p_eb.font.name = FONT_NAME
        p_eb.font.size = Pt(10.5)
        p_eb.font.color.rgb = COLOR_TEXT

# Generate all 10 SQL Slides
sql_titles = {
    24: "EDA with SQL: Unique Launch Site Names",
    25: "EDA with SQL: First 5 Records at Cape Canaveral",
    26: "EDA with SQL: Total Payload Mass Carried by NASA (CRS)",
    27: "EDA with SQL: Average Payload Mass by Booster F9 v1.1",
    28: "EDA with SQL: First Successful Ground Landing Date",
    29: "EDA with SQL: Successful Drone Ship Landings (4000-6000 kg)",
    30: "EDA with SQL: Total Successful and Failure Mission Outcomes",
    31: "EDA with SQL: Boosters Carrying Maximum Payload Mass",
    32: "EDA with SQL: Failed Drone Ship Landing Outcomes in 2015",
    33: "EDA with SQL: Ranked Count of Landing Outcomes (2010–2017)"
}

for s_num, t_text in sql_titles.items():
    build_sql_slide(prs.slides[s_num - 1], s_num, t_text)

# SLIDE 35: Folium Map 1 (Global Launch Sites)
s35 = prs.slides[34]
set_title(s35, "Folium Geospatial Analytics: Global Launch Sites Distribution")
add_watermark(s35)
for sh in s35.shapes:
    if sh != s35.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Geospatial Launch Complex Placement", [
                "Coastal Proximity: All launch complexes are located directly adjacent to oceans (Atlantic in Florida, Pacific in California) to ensure jettisoned stages and flight aborts fall safely into open water.",
                "Equatorial Rotation Advantage: Cape Canaveral sites (Lat ~28.5°N) exploit Earth's rotational speed (~1,470 km/h) to provide an eastward velocity boost for equatorial orbits.",
                "Polar Trajectory Corridor: VAFB SLC-4E (Lat ~34.6°N) provides a clear southbound flight corridor over the Pacific Ocean for polar and sun-synchronous orbits without overflying populated land."
            ])
        ], font_size=11)
s35.shapes.add_picture('images/folium_map_1_launch_sites.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 36: Folium Map 2 (Marker Clusters)
s36 = prs.slides[35]
set_title(s36, "Folium Cluster Analytics: Landing Outcomes by Launch Site")
add_watermark(s36)
for sh in s36.shapes:
    if sh != s36.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Marker Clustering Observations", [
                "Color-Coded Telemetry: Green markers designate successful first-stage recoveries; red markers designate landing failures or unattempted recoveries.",
                "KSC LC-39A Leadership: KSC LC-39A exhibits the highest concentration of successful landing clusters (10 successes out of 13 launches = 76.9%).",
                "Temporal Maturity: Early launches at CCAFS LC-40 show dense red failure clusters, while recent clusters across all active pads are overwhelmingly green."
            ])
        ], font_size=11)
s36.shapes.add_picture('images/folium_map_2_marker_clusters.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 37: Folium Map 3 (Proximities)
s37 = prs.slides[36]
set_title(s37, "Folium Proximity Analytics: Distances to Transport & Cities")
add_watermark(s37)
for sh in s37.shapes:
    if sh != s37.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Haversine Infrastructure Proximity Analysis", [
                "Coastline Buffer (0.85 km): Launch pads are situated within 1 km of the shoreline for emergency abort containment and ASDS drone ship offloading logistics.",
                "Railways & Highways (1.25–1.35 km): Titan III NASA Railroad and Samuel C. Phillips Parkway enable efficient transportation of 12-foot diameter booster stages and propellants.",
                "Urban Safety Buffer (16.20 km): Distances to nearest cities (Titusville / Cape Canaveral) exceed 16 km, shielding civilians from acoustic shockwaves, sonic booms, and explosive risks."
            ])
        ], font_size=11)
s37.shapes.add_picture('images/folium_map_3_proximities.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 39: Plotly Dash 1 (Pie Chart All Sites)
s39 = prs.slides[38]
set_title(s39, "Plotly Dash: Total Successful Launches by Launch Site")
add_watermark(s39)
for sh in s39.shapes:
    if sh != s39.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Key Observations & Site Contribution", [
                "Overall Proportions: KSC LC-39A contributed the largest share of successful first-stage landings (41.7%), followed by CCAFS SLC-40 (29.2%), VAFB SLC-4E (20.8%), and CCAFS LC-40 (8.3%).",
                "Infrastructure Efficiency: KSC LC-39A's historic pad infrastructure accommodates heavy-lift configurations and rapid turnaround for commercial and NASA missions.",
                "Operational Focus: Modern launch cadence is concentrated at KSC LC-39A and CCAFS SLC-40."
            ])
        ], font_size=11)
s39.shapes.add_picture('images/dash_pie_all_sites.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 40: Plotly Dash 2 (Pie Chart Best Site)
s40 = prs.slides[39]
set_title(s40, "Plotly Dash: Success vs. Failure Ratio for Best Launch Site")
add_watermark(s40)
for sh in s40.shapes:
    if sh != s40.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("KSC LC-39A Performance Breakdown", [
                "Highest Success Rate: KSC LC-39A achieved a stellar 76.9% landing success rate (10 successful recoveries vs. 3 failures).",
                "Advanced Pad Capabilities: Pad 39A incorporates modernized Falcon 9 launch mount systems and dedicated telemetry tracking corridors.",
                "Risk Assessment: Customers launching from KSC LC-39A benefit from the lowest booster recovery failure risk in the commercial space market."
            ])
        ], font_size=11)
s40.shapes.add_picture('images/dash_pie_best_site.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 41: Plotly Dash 3 (Scatter Payload vs Outcome)
s41 = prs.slides[40]
set_title(s41, "Plotly Dash: Payload Mass vs. Launch Outcome by Booster Version")
add_watermark(s41)
for sh in s41.shapes:
    if sh != s41.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Key Observations & Range Slider Dynamics", [
                "Sweet-Spot Payload Range: The highest concentration of successful landings occurs between 2,000 kg and 5,500 kg across all booster types.",
                "Booster Version Impact: Early v1.0 and v1.1 boosters suffered high failure rates across all payload masses; Block 5 (B5) and Full Thrust (FT) boosters demonstrate consistent landing success even for heavy payloads (>9,000 kg).",
                "Starlink Payload Scaling: Missions carrying 15,600 kg achieve high landing success rates when deployed to Low Earth Orbit via drone ship recoveries."
            ])
        ], font_size=11)
s41.shapes.add_picture('images/dash_scatter_payload_outcome.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 43: Model Accuracy Comparison
s43 = prs.slides[42]
set_title(s43, "Classification Model Performance Comparison")
add_watermark(s43)
for sh in s43.shapes:
    if sh != s43.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Model Performance & Cross-Validation Results", [
                "Benchmark Accuracy: All 4 classification models (Logistic Regression, SVM, Decision Tree, and KNN) achieved an identical 83.33% Test Accuracy on the 18-sample unseen test set.",
                "Train CV Scores (10-Fold):",
                "  • Decision Tree: Highest CV score of 87.14% (criterion='gini', max_depth=4, min_samples_leaf=2).",
                "  • Support Vector Machine: 84.82% CV score (kernel='sigmoid', C=1.0, gamma=0.0316).",
                "  • K-Nearest Neighbors: 84.82% CV score (n_neighbors=10, p=1).",
                "  • Logistic Regression: 84.64% CV score (C=0.01, penalty='l2').",
                "Model Selection: Decision Tree and Logistic Regression offer the best operational balance of high cross-validation stability, zero overfitting, and complete explainability."
            ])
        ], font_size=10.5)
s43.shapes.add_picture('images/ml_model_accuracy_comparison.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 44: Confusion Matrix
s44 = prs.slides[43]
set_title(s44, "Confusion Matrix Analysis: Best Performing Model")
add_watermark(s44)
for sh in s44.shapes:
    if sh != s44.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(4.8)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Confusion Matrix Breakdown & Error Analysis", [
                "True Positives (TP = 12): The model correctly predicted 12 successful landing recoveries.",
                "True Negatives (TN = 3): The model correctly predicted 3 landing failures.",
                "False Positives (FP = 3): The model predicted success for 3 missions that ultimately failed.",
                "False Negatives (FN = 0): Zero successful landings were misclassified as failures!",
                "Critical Operational Value:",
                "  • 100% Recall (Sensitivity): In aerospace recovery logistics, zero false negatives ensures that recovery recovery ships/teams are never mistakenly called off when a booster can be saved.",
                "  • High Precision (80.0%): Provides commercial customers with high confidence in recovery likelihood."
            ])
        ], font_size=10.5)
s44.shapes.add_picture('images/ml_confusion_matrix.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# SLIDE 45: Conclusions
s45 = prs.slides[44]
set_title(s45, "Conclusions & Commercial Recommendations")
add_watermark(s45)
for sh in s45.shapes:
    if sh != s45.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(11.5)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Core Analytical Conclusions", [
                "Technical Mastery: SpaceX successfully scaled Falcon 9 first-stage recovery from an experimental 0% success rate in 2010–2012 to an industrial >80% success rate in 2018–2020, driven by Block 5 reusability improvements.",
                "Geographic & Orbital Sweet Spots: Launch complexes located near coastlines maximize orbital velocity while ensuring abort safety. Polar, SSO, and VLEO orbits yield near-perfect recoveries, whereas high-energy GTO flights carry higher recovery risk (~50%).",
                "Predictive Viability: Supervised machine learning algorithms achieve 83.33% test accuracy with 100% recall for landing success, confirming high predictive feasibility based on pre-launch flight parameters."
            ]),
            ("Actionable Commercial & Strategic Recommendations", [
                "Competitor Bidding Strategy: Competitor launch providers (e.g. Arianespace, ULA) can utilize this predictive framework to anticipate whether SpaceX will attempt and succeed in booster recovery on a given RFP. For GTO missions, competitors can exploit SpaceX's lower recovery margins to offer competitive pricing.",
                "Insurance & Risk Underwriting: Marine and space insurance underwriters can adjust premiums dynamically based on flight parameters (e.g., lower premiums for KSC LC-39A and LEO/VLEO missions; adjusted rates for high-mass GTO flights).",
                "Payload Optimization: Commercial satellite operators should target payload masses between 2,000 kg and 5,500 kg to maximize first-stage recovery probability and secure negotiated discounts on flight-proven boosters."
            ])
        ], font_size=11)

# SLIDE 46: Appendix
s46 = prs.slides[45]
set_title(s46, "Appendix: Data Sources, Technologies & GitHub Links")
add_watermark(s46)
for sh in s46.shapes:
    if sh != s46.shapes.title and sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(11.5)
        sh.height = Inches(5.2)
        populate_text_bullets(sh, [
            ("Verified GitHub Repository & Code Deliverables", [
                "Main Repository: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction",
                "1. Data Collection API: notebooks/1_SpaceX_Data_Collection_API.ipynb",
                "2. Data Collection Scraping: notebooks/2_SpaceX_Data_Collection_Web_Scraping.ipynb",
                "3. Data Wrangling: notebooks/3_SpaceX_Data_Wrangling.ipynb",
                "4. EDA Data Visualization: notebooks/4_SpaceX_EDA_Data_Visualization.ipynb",
                "5. EDA with SQL: notebooks/5_SpaceX_EDA_SQL.ipynb",
                "6. Folium Interactive Maps: notebooks/6_SpaceX_Interactive_Visual_Analytics_Folium.ipynb",
                "7. Plotly Dash Web Application: spacex_dash_app.py",
                "8. Machine Learning Prediction: notebooks/7_SpaceX_Machine_Learning_Prediction.ipynb"
            ]),
            ("Data Dictionary & Technology Stack", [
                "Data Sources: SpaceX REST API v4, Wikipedia Falcon 9 historical archives, IBM Skills Network database.",
                "Technologies Utilized: Python 3.10, Pandas, NumPy, Scikit-Learn, Matplotlib, Seaborn, Folium, Plotly Dash, SQLite, python-pptx."
            ])
        ], font_size=10.5)

# SLIDE 47: Thank You Slide
s47 = prs.slides[46]
set_title(s47, "Thank You / Questions & Discussion")
add_watermark(s47)
tx_box = s47.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(9.2), Inches(3.8))
tf_ty = tx_box.text_frame
tf_ty.word_wrap = True

p_ty1 = tf_ty.paragraphs[0]
p_ty1.alignment = PP_ALIGN.CENTER
p_ty1.text = "Thank You for Your Time & Evaluation"
p_ty1.font.name = FONT_NAME
p_ty1.font.size = Pt(28)
p_ty1.font.bold = True
p_ty1.font.color.rgb = COLOR_TITLE
p_ty1.space_after = Pt(16)

p_ty2 = tf_ty.add_paragraph()
p_ty2.alignment = PP_ALIGN.CENTER
p_ty2.text = "SpaceX Falcon 9 First Stage Landing Prediction Project"
p_ty2.font.name = FONT_NAME
p_ty2.font.size = Pt(18)
p_ty2.font.color.rgb = COLOR_HIGHLIGHT
p_ty2.space_after = Pt(20)

p_ty3 = tf_ty.add_paragraph()
p_ty3.alignment = PP_ALIGN.CENTER
p_ty3.text = "Author: Heet Patel\nIBM Data Science Professional Certificate | Coursera"
p_ty3.font.name = FONT_NAME
p_ty3.font.size = Pt(14)
p_ty3.font.bold = True
p_ty3.font.color.rgb = COLOR_TEXT
p_ty3.space_after = Pt(14)

p_ty4 = tf_ty.add_paragraph()
p_ty4.alignment = PP_ALIGN.CENTER
p_ty4.text = "GitHub Repository: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction"
p_ty4.font.name = FONT_NAME
p_ty4.font.size = Pt(13)
p_ty4.font.bold = True
p_ty4.font.color.rgb = COLOR_ACCENT

# Save output presentation
output_pptx = "Data Science Capstone Project Report.pptx"
prs.save(output_pptx)
# Also save over the template so both exist and are identical
prs.save("ds-capstone-template-coursera.pptx")
print(f"Successfully generated and saved: {output_pptx} and ds-capstone-template-coursera.pptx")
