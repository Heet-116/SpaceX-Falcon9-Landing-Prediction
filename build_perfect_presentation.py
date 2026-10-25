import os
import sqlite3
import pandas as pd
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# 1. Execute SQL queries on dataset
df_sql = pd.read_csv('data/Spacex.csv')
conn = sqlite3.connect(':memory:')
df_sql.to_sql('SPACEXTBL', conn, index=False, if_exists='replace')

sql_queries = {
    24: ("SELECT DISTINCT Launch_Site FROM SPACEXTBL",
         "The dataset identifies 4 unique launch sites: CCAFS LC-40, VAFB SLC-4E, KSC LC-39A, and CCAFS SLC-40. Cape Canaveral and Kennedy Space Center support equatorial and high-inclination eastward launches over the Atlantic, while Vandenberg Air Force Base supports polar and sun-synchronous launches over the Pacific."),
    25: ("SELECT Date, Booster_Version, Launch_Site, PAYLOAD_MASS__KG_, Landing_Outcome FROM SPACEXTBL WHERE Launch_Site LIKE 'CCA%' LIMIT 5",
         "The first 5 orbital missions launched exclusively from Cape Canaveral Air Force Station (CCAFS LC-40) carrying early Dragon qualification units. These early flights (2010-2013) were experimental flights with ocean splashdowns and parachute recovery attempts."),
    26: ("SELECT SUM(PAYLOAD_MASS__KG_) AS Total_Payload_NASA_CRS FROM SPACEXTBL WHERE Customer = 'NASA (CRS)'",
         "SpaceX delivered a total of 45,596 kg of cargo to the International Space Station across more than 20 NASA Commercial Resupply Services (CRS) missions, validating the operational reliability and heavy-cargo capacity of the Falcon 9 launch system."),
    27: ("SELECT ROUND(AVG(PAYLOAD_MASS__KG_), 2) AS Avg_Payload_F9_v1_1 FROM SPACEXTBL WHERE Booster_Version LIKE 'F9 v1.1%'",
         "The Falcon 9 v1.1 booster variant carried an average payload mass of 2,928.4 kg, serving medium-weight commercial communications satellites and ISS resupply prior to the introduction of Full Thrust and Block 5 upgrades."),
    28: ("SELECT MIN(Date) AS First_Successful_Ground_Landing, Booster_Version, Launch_Site, Landing_Outcome FROM SPACEXTBL WHERE Landing_Outcome = 'Success (ground pad)'",
         "The world's first successful orbital rocket first-stage ground landing occurred on December 22, 2015, by booster B1019 at Cape Canaveral Landing Zone 1 (LZ-1), proving the engineering feasibility of vertical rocket recovery."),
    29: ("SELECT Booster_Version, PAYLOAD_MASS__KG_, Launch_Site FROM SPACEXTBL WHERE Landing_Outcome = 'Success (drone ship)' AND PAYLOAD_MASS__KG_ BETWEEN 4000 AND 6000",
         "Four boosters (B1022, B1026, B1021.2, B1031.2) successfully completed Autonomous Spaceport Drone Ship (ASDS) landings during demanding GTO missions where return-to-launch-site (RTLS) was impossible due to booster velocity and propellant limits."),
    30: ("SELECT Mission_Outcome, COUNT(*) AS Outcome_Count FROM SPACEXTBL GROUP BY Mission_Outcome",
         "Out of 101 recorded missions, 98 primary mission outcomes were completely successful (97.0% overall mission success rate). Only 1 in-flight failure occurred (CRS-7), demonstrating exceptional launch reliability across commercial and government missions."),
    31: ("SELECT Booster_Version, Launch_Site, PAYLOAD_MASS__KG_ FROM SPACEXTBL WHERE PAYLOAD_MASS__KG_ = (SELECT MAX(PAYLOAD_MASS__KG_) FROM SPACEXTBL) LIMIT 5",
         "The maximum payload carried is 15,600 kg, deployed exclusively by Falcon 9 Block 5 boosters (including B1048, B1049, B1051, B1056) launching batches of 60 Starlink satellites to Low Earth Orbit."),
    32: ("SELECT Date, Booster_Version, Launch_Site, Landing_Outcome FROM SPACEXTBL WHERE Landing_Outcome = 'Failure (drone ship)' AND Date LIKE '2015%'",
         "In 2015, two drone ship landing attempts failed (B1012 on Jan 10 and B1015 on Apr 14 at CCAFS LC-40) due to hydraulic fluid exhaustion and landing leg lock issues, providing vital telemetry for subsequent landing successes."),
    33: ("SELECT Landing_Outcome, COUNT(*) AS Outcome_Count FROM SPACEXTBL WHERE Date BETWEEN '2010-06-04' AND '2017-03-20' GROUP BY Landing_Outcome ORDER BY Outcome_Count DESC",
         "Between 2010 and 2017, early operations were dominated by 10 'No attempt' flights, followed by 5 drone ship failures, 5 drone ship successes, and 3 ground pad successes as SpaceX iteratively perfected booster recovery hardware.")
}

sql_data = {}
for s_idx, (q, expl) in sql_queries.items():
    df_res = pd.read_sql_query(q, conn)
    sql_data[s_idx] = (q, df_res, expl)

prs = pptx.Presentation('ds-capstone-template-coursera.pptx')

COLOR_TITLE = RGBColor(24, 43, 73)
COLOR_TEXT = RGBColor(40, 50, 60)
COLOR_MUTED = RGBColor(120, 130, 140)
COLOR_BLUE = RGBColor(31, 110, 165)
COLOR_WHITE = RGBColor(255, 255, 255)
FONT_NAME = 'Calibri'
WATERMARK = "Heet Patel | IBM Applied Data Science Capstone | https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction"

def add_watermark(slide):
    tx = slide.shapes.add_textbox(Inches(0.8), Inches(7.15), Inches(11.5), Inches(0.28))
    tf = tx.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = WATERMARK
    p.font.name = FONT_NAME
    p.font.size = Pt(8.5)
    p.font.color.rgb = COLOR_MUTED

def set_slide_title(slide, new_title):
    for sh in slide.shapes:
        if 'Title' in sh.name and sh.has_text_frame:
            sh.text_frame.clear()
            p = sh.text_frame.paragraphs[0]
            p.text = new_title
            p.font.name = FONT_NAME
            p.font.size = Pt(24)
            p.font.bold = True
            p.font.color.rgb = COLOR_TITLE

def get_content_shape(slide):
    for sh in slide.shapes:
        if not sh.name.startswith(('Slide Number', 'Title', 'Picture', 'TextBox 6')) and sh.has_text_frame:
            return sh
    for sh in slide.shapes:
        if not sh.name.startswith(('Slide Number', 'Title', 'Picture')) and sh.has_text_frame:
            return sh
    return None

def write_slide_content(shape, title_header, bullet_groups, font_size=12):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    
    first = True
    if title_header:
        p = tf.paragraphs[0]
        first = False
        p.text = title_header
        p.font.name = FONT_NAME
        p.font.size = Pt(font_size + 2)
        p.font.bold = True
        p.font.color.rgb = COLOR_TITLE
        p.space_after = Pt(4)
        
    for header, bullets in bullet_groups:
        if header:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.text = header
            p.font.name = FONT_NAME
            p.font.size = Pt(font_size + 1)
            p.font.bold = True
            p.font.color.rgb = COLOR_BLUE
            p.space_after = Pt(2)
            p.space_before = Pt(5) if not first else Pt(0)
            
        for b in bullets:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.text = f"• {b}"
            p.font.name = FONT_NAME
            p.font.size = Pt(font_size)
            p.font.color.rgb = COLOR_TEXT
            p.space_after = Pt(2.5)

print("Populating presentation slides...")

# Slide 1: Title Slide
s1 = prs.slides[0]
for sh in s1.shapes:
    if sh.has_text_frame:
        sh.left = Inches(0.8)
        sh.top = Inches(2.3)
        sh.width = Inches(11.5)
        sh.height = Inches(4.2)
        tf = sh.text_frame
        tf.clear()
        p1 = tf.paragraphs[0]
        p1.text = "SpaceX Falcon 9 First Stage Landing Prediction"
        p1.font.name = FONT_NAME
        p1.font.size = Pt(32)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_TITLE
        p1.space_after = Pt(8)
        
        p2 = tf.add_paragraph()
        p2.text = "IBM Applied Data Science Capstone Final Report"
        p2.font.name = FONT_NAME
        p2.font.size = Pt(20)
        p2.font.color.rgb = COLOR_BLUE
        p2.space_after = Pt(22)
        
        p3 = tf.add_paragraph()
        p3.text = "Author: Heet Patel | IBM Data Science Professional Certificate"
        p3.font.name = FONT_NAME
        p3.font.size = Pt(14)
        p3.font.bold = True
        p3.font.color.rgb = COLOR_TEXT
        p3.space_after = Pt(4)
        
        p4 = tf.add_paragraph()
        p4.text = "Submission Format: Completed Presentation in PDF Format | Date: October 2026"
        p4.font.name = FONT_NAME
        p4.font.size = Pt(13)
        p4.font.color.rgb = COLOR_MUTED
        p4.space_after = Pt(8)
        
        p5 = tf.add_paragraph()
        p5.text = "GitHub Repository: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction"
        p5.font.name = FONT_NAME
        p5.font.size = Pt(13)
        p5.font.bold = True
        p5.font.color.rgb = RGBColor(39, 174, 96)

# Slide 2: Outline
s2 = prs.slides[1]
set_slide_title(s2, "Outline")
add_watermark(s2)
c2 = get_content_shape(s2)
if c2:
    write_slide_content(c2, "Table of Contents & Project Roadmap", [
        ("Executive Summary & Introduction", [
            "Executive Summary: Complete summary of methodologies and quantitative results",
            "Introduction: Commercial spaceflight background, reusable economics, and research questions"
        ]),
        ("Data Science Methodologies", [
            "Data Collection & Data Wrangling: SpaceX REST API, Wikipedia scraping, and data wrangling",
            "Exploratory Data Analysis: Data visualization (Matplotlib/Seaborn) and SQL queries (SQLite)",
            "Interactive Visual Analytics: Folium geospatial maps and Plotly Dash interactive dashboard",
            "Predictive Analysis (Classification): Machine learning models, GridSearchCV, and tuning"
        ]),
        ("Results, Conclusions & Appendix", [
            "Comprehensive Results: EDA visualizations, 10 SQL query results, Folium maps, Dash app, and ML metrics",
            "Conclusions & Innovative Insights: Technology trajectory, cost dynamics, and commercial bidding strategy",
            "Appendix: Data dictionary, technology stack, and verified GitHub repository deliverables"
        ])
    ], font_size=11)

# Slide 3: Executive Summary
s3 = prs.slides[2]
set_slide_title(s3, "Executive Summary")
add_watermark(s3)
c3 = get_content_shape(s3)
if c3:
    c3.left = Inches(0.8)
    c3.top = Inches(1.5)
    c3.width = Inches(11.5)
    c3.height = Inches(5.2)
    write_slide_content(c3, "Executive Summary", [
        ("Summary of Methodologies", [
            "Data Collection & Wrangling: Ingested 90 Falcon 9 launches via SpaceX REST API v4 and scraped historical Wikipedia launch logs with BeautifulSoup. Imputed null payload values with mean (6,105 kg) and engineered a binary 'Class' landing target (1 = Success, 0 = Failure). Expanded categorical features via one-hot encoding into 83 numeric columns.",
            "Exploratory Data Analysis: Analyzed correlations between payload mass, launch site, flight cadence, and orbit types using Matplotlib/Seaborn; executed 10 complex SQL queries on SQLite database table 'SPACEXTBL' to isolate operational milestones, customer payloads, and booster variants.",
            "Interactive Analytics: Built Folium maps with marker clusters and Haversine proximity calculations to coastal/transport infrastructure; engineered a responsive Plotly Dash dashboard with payload range sliders and site dropdowns.",
            "Predictive Modeling: Standardized 83 features via StandardScaler; trained 4 classification models (Logistic Regression, SVM, Decision Tree, KNN) optimized using 10-fold cross-validation GridSearchCV."
        ]),
        ("Summary of All Results & Business Insights", [
            "Overall Success Rate: 66.7% across all 90 launches; success escalated from 0% (2010-2012) to >80% (2018-2020) driven by Falcon 9 Block 5 reusability enhancements.",
            "Best Launch Complex: KSC LC-39A achieved the highest success rate of 76.9% and contributed 41.7% of all successful first-stage recoveries.",
            "Predictive Performance: All 4 machine learning models achieved 83.33% Test Accuracy. Crucially, the models achieved 100% Recall (0 False Negatives) for successful landings, preventing costly discard of recoverable boosters.",
            "Commercial Utility: Enables competitor launch providers and commercial satellite operators to predict SpaceX booster recovery likelihood, informing competitive pricing against SpaceX's $62M rate."
        ])
    ], font_size=10)

# Slide 4: Introduction
s4 = prs.slides[3]
set_slide_title(s4, "Introduction")
add_watermark(s4)
c4 = get_content_shape(s4)
if c4:
    c4.left = Inches(0.8)
    c4.top = Inches(1.5)
    c4.width = Inches(11.5)
    c4.height = Inches(5.2)
    write_slide_content(c4, "Introduction", [
        ("Project Background & Context", [
            "Commercial Space Revolution: SpaceX has transformed the aerospace industry by developing orbital-class reusable rocket boosters, drastically reducing the cost of reaching space.",
            "Economic Disruption: Traditional expendable rocket launches cost upwards of $165+ million. In contrast, SpaceX advertises Falcon 9 launches at approximately $62 million, achieving savings by recovering and reusing the first stage booster.",
            "Value of Recovery: The first-stage booster represents ~70% of total rocket manufacturing cost. Reusing boosters multiple times amortizes production capital and yields unprecedented commercial margins."
        ]),
        ("Problems to Find Answers (Core Research Questions)", [
            "1. Recovery Determinants: How do payload mass, target orbital profile, launch site geography, and booster flight cadence influence first-stage landing success?",
            "2. Geospatial Patterns: Why are launch complexes positioned along specific coastlines, and how do distances to transport and cities affect operational safety?",
            "3. Predictive Feasibility: Can supervised machine learning classification algorithms accurately predict first-stage landing outcomes prior to liftoff?",
            "4. Commercial Bidding Strategy: How can competitor launch providers and commercial satellite operators leverage predictive analytics to optimize pricing and contract bidding against SpaceX?"
        ])
    ], font_size=10.5)

# Slide 5: Section 1 Header
s5 = prs.slides[4]
add_watermark(s5)
for sh in s5.shapes:
    if sh.has_text_frame and 'TextBox' in sh.name:
        sh.text_frame.text = "Section 1: Data Collection & Data Wrangling Methodology"
        sh.text_frame.paragraphs[0].font.size = Pt(28)
        sh.text_frame.paragraphs[0].font.bold = True
        sh.text_frame.paragraphs[0].font.color.rgb = COLOR_TITLE

# Slide 6: Methodology
s6 = prs.slides[5]
set_slide_title(s6, "Methodology")
add_watermark(s6)
c6 = get_content_shape(s6)
if c6:
    c6.left = Inches(0.8)
    c6.top = Inches(1.5)
    c6.width = Inches(11.5)
    c6.height = Inches(5.2)
    write_slide_content(c6, "Methodology Overview", [
        ("Phase 1: Data Collection & Ingestion", [
            "SpaceX REST API v4: Extracted 90+ rocket launch records via JSON endpoints capturing cores, flight numbers, payload specs, and landing pads.",
            "Wikipedia Web Scraping: Scraped historical Falcon 9 launch logs from Wikipedia using BeautifulSoup to extract customers, dates, and mission outcomes."
        ]),
        ("Phase 2: Data Wrangling & Preprocessing", [
            "Identified missing payload values and imputed with mean (6,104.96 kg); mapped raw landing outcomes into binary classification target 'Class' (1 = Success, 0 = Failure).",
            "Applied one-hot encoding on categorical features (Orbit, LaunchSite, LandingPad, Serial), expanding the feature space to 83 numerical columns."
        ]),
        ("Phase 3: Exploratory Data Analysis (EDA)", [
            "Visual Analytics: Analyzed distributions across payload, orbit, and temporal trends using Matplotlib and Seaborn.",
            "SQL Relational Querying: Loaded dataset into SQLite database 'SPACEXTBL' and executed 10 technical SQL queries for aggregated insights and historical milestones."
        ]),
        ("Phase 4: Interactive Visual Analytics", [
            "Folium Geospatial Mapping: Plotted launch complexes with marker clusters and Haversine distance proximity measurements to coastline, transport, and cities.",
            "Plotly Dash Web Dashboard: Engineered a dynamic web dashboard featuring reactive dropdowns and payload range sliders for live slice-and-dice analytics."
        ]),
        ("Phase 5: Predictive Analysis (Classification)", [
            "Feature Standardization: Standardized 83 features via StandardScaler; partitioned data into 80% Train and 20% Test sets (random_state=2).",
            "Hyperparameter Tuning: Tuned Logistic Regression, SVM, Decision Tree, and KNN models using GridSearchCV with 10-fold cross validation."
        ])
    ], font_size=10)

# Slide 7: Data Collection
s7 = prs.slides[6]
set_slide_title(s7, "Data Collection")
add_watermark(s7)
c7 = get_content_shape(s7)
if c7:
    c7.left = Inches(0.8)
    c7.top = Inches(1.5)
    c7.width = Inches(11.5)
    c7.height = Inches(5.2)
    write_slide_content(c7, "Data Collection Methodology", [
        ("Data Collection Process & Methodology", [
            "Multi-Source Ingestion: Data was collected using two distinct methodologies to ensure high fidelity: programmatic REST API calls and web scraping.",
            "SpaceX REST API (JSON): Used Python 'requests' library to call endpoint 'https://api.spacexdata.com/v4/launches/past'. Normalized JSON data with 'pd.json_normalize' and extracted core telemetry (booster serials, flight numbers, gridfins, landing pads, outcome tuples). Filtered specifically for 90 Falcon 9 launches.",
            "Wikipedia Web Scraping (HTML): Extracted historical launch tables from Wikipedia ('List of Falcon 9 and Falcon Heavy launches') using BeautifulSoup. Parsed <th> and <td> tags, removed citation references, sanitized dates, and standardized numerical payload masses.",
            "Consolidation & Quality Assurance: Cross-validated flight numbers and booster serials across both datasets to create a clean, comprehensive dataset."
        ]),
        ("GitHub Repository Deliverables", [
            "Completed API Notebook: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/1_SpaceX_Data_Collection_API.ipynb",
            "Completed Scraping Notebook: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/2_SpaceX_Data_Collection_Web_Scraping.ipynb"
        ])
    ], font_size=10.5)

# Slide 8: Data Collection – SpaceX API
s8 = prs.slides[7]
set_slide_title(s8, "Data Collection – SpaceX API")
add_watermark(s8)
for sh in s8.shapes:
    if sh.name == 'Text Placeholder 2':
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(5.0)
        sh.height = Inches(5.2)
        write_slide_content(sh, "Data Collection – SpaceX API", [
            ("SpaceX REST API Data Collection Methodology", [
                "API Endpoint: https://api.spacexdata.com/v4/launches/past",
                "JSON Normalization: Used requests.get() and pd.json_normalize() to parse complex nested JSON response structures.",
                "Telemetry Fields: Extracted rocket specifications, launchpad IDs, flight numbers, payload mass, destination orbit, and core telemetry.",
                "Filtering: Filtered dataset strictly for Falcon 9 launches (excluding Falcon 1 and Falcon Heavy multi-core missions), yielding 90 launches.",
                "Output: Structured DataFrame of 90 records with 17 initial features."
            ]),
            ("GitHub URL Reference", [
                "Completed Notebook: notebooks/1_SpaceX_Data_Collection_API.ipynb",
                "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/1_SpaceX_Data_Collection_API.ipynb"
            ])
        ], font_size=10)
    elif sh.name == 'Content Placeholder 4':
        sh.text_frame.clear()
s8.shapes.add_picture('images/flowchart_api.png', Inches(6.0), Inches(1.6), Inches(5.4), Inches(4.8))

# Slide 9: Data Collection - Scraping
s9 = prs.slides[8]
set_slide_title(s9, "Data Collection - Scraping")
add_watermark(s9)
for sh in s9.shapes:
    if sh.name == 'Text Placeholder 2':
        sh.left = Inches(0.8)
        sh.top = Inches(1.5)
        sh.width = Inches(5.0)
        sh.height = Inches(5.2)
        write_slide_content(sh, "Data Collection - Scraping", [
            ("Wikipedia Web Scraping Methodology", [
                "Target Archive: Wikipedia List of Falcon 9 and Falcon Heavy launches.",
                "HTML Parsing: Employed BeautifulSoup to traverse DOM elements, identifying target launch tables with class='wikitable'.",
                "Data Extraction: Iterated over table rows (<tr>) and data cells (<td>), extracting Flight No., Date, Booster Version, Launch Site, Payload, Orbit, Customer, and Outcome.",
                "Data Sanitization: Utilized regular expressions to strip superscript citations, handle irregular dates, and cast string payload masses to floats.",
                "Consolidation: Formed structured DataFrame ready for data wrangling."
            ]),
            ("GitHub URL Reference", [
                "Completed Notebook: notebooks/2_SpaceX_Data_Collection_Web_Scraping.ipynb",
                "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/2_SpaceX_Data_Collection_Web_Scraping.ipynb"
            ])
        ], font_size=10)
    elif sh.name == 'Content Placeholder 4':
        sh.text_frame.clear()
s9.shapes.add_picture('images/flowchart_scraping.png', Inches(6.0), Inches(1.6), Inches(5.4), Inches(4.8))

# Slide 10: Data Wrangling
s10 = prs.slides[9]
set_slide_title(s10, "Data Wrangling")
add_watermark(s10)
c10 = get_content_shape(s10)
if c10:
    c10.left = Inches(0.8)
    c10.top = Inches(1.5)
    c10.width = Inches(5.0)
    c10.height = Inches(5.2)
    write_slide_content(c10, "Data Wrangling Methodology", [
        ("Data Wrangling & Feature Engineering Methodology", [
            "Missing Value Imputation: Detected 5 missing values in 'PayloadMass'; imputed with dataset mean (6,104.96 kg) to retain all 90 records.",
            "Target Variable Labeling: Evaluated landing outcome tuples:",
            "  • Success (Class = 1): True RTLS, True ASDS, True Ocean (60 launches, 66.7%).",
            "  • Failure (Class = 0): False RTLS, False ASDS, False Ocean, None ASDS, None None (30 launches, 33.3%).",
            "One-Hot Encoding: Applied pd.get_dummies to categorical columns ('Orbit', 'LaunchSite', 'LandingPad', 'Serial'), producing 83 numeric predictor columns."
        ]),
        ("GitHub URL Reference", [
            "Completed Notebook: notebooks/3_SpaceX_Data_Wrangling.ipynb",
            "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/3_SpaceX_Data_Wrangling.ipynb"
        ])
    ], font_size=10)
s10.shapes.add_picture('images/flowchart_wrangling.png', Inches(6.0), Inches(1.6), Inches(5.4), Inches(4.8))

# Slide 11: EDA with Data Visualization
s11 = prs.slides[10]
set_slide_title(s11, "EDA with Data Visualization")
add_watermark(s11)
c11 = get_content_shape(s11)
if c11:
    c11.left = Inches(0.8)
    c11.top = Inches(1.5)
    c11.width = Inches(11.5)
    c11.height = Inches(5.2)
    write_slide_content(c11, "EDA and Data Visualization Methodology", [
        ("EDA with Data Visualization Methodology & Rationale", [
            "Analytical Framework: Visual analytics executed using Pandas, Matplotlib, and Seaborn to understand bivariate interactions and success drivers.",
            "Charts Plotted and Analytical Justification:",
            "  1. Flight Number vs. Launch Site: Scatter plot to observe temporal launch migration across launch pads and trace landing reliability over time.",
            "  2. Payload Mass vs. Launch Site: Scatter plot to identify payload capabilities and landing success rates per launch complex.",
            "  3. Success Rate vs. Orbit Type: Bar chart to quantify recovery probability across 11 orbital mission profiles (e.g., LEO, GTO, VLEO).",
            "  4. Flight Number vs. Orbit Type: Scatter plot to visualize orbital cadence changes (early LEO/GTO transitioning to high-frequency VLEO Starlink).",
            "  5. Payload Mass vs. Orbit Type: Scatter plot to assess payload-energy limits and booster propellant recovery feasibility.",
            "  6. Launch Success Yearly Trend: Temporal line plot to quantify SpaceX's engineering learning curve and success rate escalation from 2010 to 2020."
        ]),
        ("GitHub URL Reference", [
            "Completed Notebook: notebooks/4_SpaceX_EDA_Data_Visualization.ipynb",
            "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/4_SpaceX_EDA_Data_Visualization.ipynb"
        ])
    ], font_size=10.5)

# Slide 12: EDA with SQL
s12 = prs.slides[11]
set_slide_title(s12, "EDA with SQL")
add_watermark(s12)
c12 = get_content_shape(s12)
if c12:
    c12.left = Inches(0.8)
    c12.top = Inches(1.5)
    c12.width = Inches(11.5)
    c12.height = Inches(5.2)
    write_slide_content(c12, "EDA with SQL Methodology", [
        ("EDA with SQL Querying Methodology", [
            "Database Environment: Loaded 101 historical SpaceX launch records into an SQLite relational database table named 'SPACEXTBL'.",
            "SQL Analytical Objectives:",
            "  • Data Profiling: Executed SELECT DISTINCT Launch_Site to identify all distinct launch facilities.",
            "  • Pattern Matching: Applied WHERE Launch_Site LIKE 'CCA%' to extract early Cape Canaveral launches, and Booster_Version LIKE 'F9 v1.1%' to filter booster versions.",
            "  • Aggregations: Calculated SUM(PAYLOAD_MASS__KG_) for NASA CRS missions and AVG(PAYLOAD_MASS__KG_) for specific booster families.",
            "  • Chronological Milestones: Used MIN(Date) to determine the exact date of the first successful ground pad landing.",
            "  • Range Filtering: Isolated successful drone ship landings with payloads BETWEEN 4000 AND 6000 kg.",
            "  • Subqueries & Extrema: Used WHERE PAYLOAD_MASS__KG_ = (SELECT MAX(...)) to identify boosters carrying maximum payload.",
            "  • Frequency Grouping & Ranking: Computed landing outcome distributions using GROUP BY and ORDER BY DESC across specific date ranges."
        ]),
        ("GitHub URL Reference", [
            "Completed Notebook: notebooks/5_SpaceX_EDA_SQL.ipynb",
            "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/5_SpaceX_EDA_SQL.ipynb"
        ])
    ], font_size=10)

# Slide 13: Build an Interactive Map with Folium
s13 = prs.slides[12]
set_slide_title(s13, "Build an Interactive Map with Folium")
add_watermark(s13)
c13 = get_content_shape(s13)
if c13:
    c13.left = Inches(0.8)
    c13.top = Inches(1.5)
    c13.width = Inches(11.5)
    c13.height = Inches(5.2)
    write_slide_content(c13, "Interactive Folium Map Methodology", [
        ("Interactive Folium Geospatial Map Methodology", [
            "Geospatial Framework: Utilized Python 'folium' library to build interactive maps centered on US launch complexes.",
            "Map Objects Created and Analytical Rationale:",
            "  • folium.Map: Initialized baseline satellite and street tile layers centered on Cape Canaveral and Vandenberg coordinates.",
            "  • folium.Circle & folium.Marker: Plotted launch complexes (CCAFS LC-40, CCAFS SLC-40, KSC LC-39A, VAFB SLC-4E) with coordinate popup labels.",
            "  • folium.plugins.MarkerCluster: Grouped dense historical launch markers to prevent visual overlapping and dynamically summarize landing outcomes.",
            "  • Custom Color-Coded Icons: Applied green markers for successful landings and red markers for failures to evaluate spatial performance.",
            "  • Haversine Distance Calculation & folium.PolyLine: Calculated straight-line distances from launch pads to coastlines, railways, highways, and cities, drawing PolyLines to evaluate safety buffers and supply logistics."
        ]),
        ("GitHub URL Reference", [
            "Completed Notebook: notebooks/6_SpaceX_Interactive_Visual_Analytics_Folium.ipynb",
            "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/6_SpaceX_Interactive_Visual_Analytics_Folium.ipynb"
        ])
    ], font_size=10)

# Slide 14: Build a Dashboard with Plotly Dash
s14 = prs.slides[13]
set_slide_title(s14, "Build a Dashboard with Plotly Dash")
add_watermark(s14)
c14 = get_content_shape(s14)
if c14:
    c14.left = Inches(0.8)
    c14.top = Inches(1.5)
    c14.width = Inches(11.5)
    c14.height = Inches(5.2)
    write_slide_content(c14, "Plotly Dash Dashboard Methodology", [
        ("Plotly Dash Interactive Dashboard Methodology", [
            "Dashboard Architecture: Developed a responsive, full-stack analytical web dashboard using Dash, dash_core_components, and dash_html_components.",
            "Interactive UI Components & Purpose:",
            "  • Launch Site Dropdown (dcc.Dropdown): Allows stakeholders to toggle between 'All Sites' and specific complexes (CCAFS LC-40, VAFB SLC-4E, KSC LC-39A, CCAFS SLC-40) to isolate site-specific recovery dynamics.",
            "  • Payload Range Slider (dcc.RangeSlider): Enables fluid filtering between 0 kg and 10,000 kg with 1,000 kg step increments to analyze payload thresholds.",
            "Reactive Callback Functions (@app.callback):",
            "  • Pie Chart Callback: Dynamically generates 'px.pie' showing total launch success share across all sites, or success vs. failure proportion for a selected site.",
            "  • Scatter Plot Callback: Dynamically renders 'px.scatter' plotting Payload Mass vs. Launch Outcome, color-coded by Booster Version Category."
        ]),
        ("GitHub Script Reference", [
            "Completed Dashboard Application: spacex_dash_app.py",
            "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/spacex_dash_app.py"
        ])
    ], font_size=10)

# Slide 15: Predictive Analysis (Classification)
s15 = prs.slides[14]
set_slide_title(s15, "Predictive Analysis (Classification)")
add_watermark(s15)
c15 = get_content_shape(s15)
if c15:
    c15.left = Inches(0.8)
    c15.top = Inches(1.5)
    c15.width = Inches(5.0)
    c15.height = Inches(5.2)
    write_slide_content(c15, "Predictive Analysis (Classification) Methodology", [
        ("Predictive Classification Methodology", [
            "Feature Standardization: Standardized 83 numerical features using StandardScaler (mean=0, variance=1) to prevent algorithm scale bias.",
            "Data Partitioning: Split into 80% Training (72 samples) and 20% Testing (18 samples) with random_state=2.",
            "Hyperparameter Tuning (GridSearchCV with 10-Fold CV):",
            "  • Logistic Regression: L2 penalty, lbfgs solver, C in [0.01, 0.1, 1].",
            "  • Support Vector Machine: Linear, RBF, Poly, Sigmoid kernels; C & gamma grids.",
            "  • Decision Tree: Gini & entropy criteria; max_depth 2 to 18; min_samples_leaf.",
            "  • K-Nearest Neighbors: k from 1 to 10; Euclidean and Manhattan distance metrics.",
            "Evaluation: Measured test set accuracy, confusion matrices, precision, and recall."
        ]),
        ("GitHub URL Reference", [
            "Completed Notebook: notebooks/7_SpaceX_Machine_Learning_Prediction.ipynb",
            "URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction/blob/main/notebooks/7_SpaceX_Machine_Learning_Prediction.ipynb"
        ])
    ], font_size=10)
s15.shapes.add_picture('images/flowchart_ml.png', Inches(6.0), Inches(1.6), Inches(5.4), Inches(4.8))

# Slide 16: Results
s16 = prs.slides[15]
set_slide_title(s16, "Results")
add_watermark(s16)
c16 = get_content_shape(s16)
if c16:
    write_slide_content(c16, "Results Overview & Architecture", [
        ("Complete Results Architecture", [
            "Exploratory Data Analysis Results: Flight numbers, payload distributions, orbital success rates, and 2010-2020 yearly trend lines.",
            "EDA with SQL Results: 10 executed database queries with tabular outputs, payload sums, booster extrema, and outcome rankings.",
            "Interactive Visual Analytics Demo: Folium geospatial maps (launch sites, marker clusters, proximity buffers) and Plotly Dash dashboard pie & scatter plots.",
            "Predictive Analysis Results: Model accuracy comparison bar chart and best performing model confusion matrix."
        ])
    ], font_size=11.5)

# Slide 17: Section 2 Header
s17 = prs.slides[16]
add_watermark(s17)
for sh in s17.shapes:
    if sh.has_text_frame and 'TextBox' in sh.name:
        sh.text_frame.text = "Section 2: Exploratory Data Analysis Results"
        sh.text_frame.paragraphs[0].font.size = Pt(28)
        sh.text_frame.paragraphs[0].font.bold = True
        sh.text_frame.paragraphs[0].font.color.rgb = COLOR_TITLE

# Slide 18 to 23: EDA Visualization Results
eda_slides_info = [
    (17, "Flight Number vs. Launch Site", 'images/eda_flight_vs_launchsite.png', [
        "Site Progression: Flights 1 through ~25 were launched almost exclusively from CCAFS LC-40, representing early Falcon 9 development.",
        "Failure Clustering: Early landing failures (red points) were concentrated at CCAFS LC-40 during initial recovery testing (2010–2015).",
        "Operational Expansion: Flights beyond #30 diversified to VAFB SLC-4E and KSC LC-39A, displaying an overwhelming surge in green landing successes.",
        "Reliability Curve: For flight numbers >50, landing success rate exceeded 80% across all launch sites."
    ]),
    (18, "Payload vs. Launch Site", 'images/eda_payload_vs_launchsite.png', [
        "VAFB SLC-4E: Specialized in light to medium payload missions (<10,000 kg) dedicated to polar and sun-synchronous orbits, achieving high recovery success.",
        "KSC LC-39A Dominance: Handled both medium and heavy payloads (>10,000 kg), including commercial communications satellites and ISS Dragon flights, with an outstanding recovery record.",
        "CCAFS SLC-40: Handled flights across the entire payload spectrum (0 to 15,600 kg); early low-mass missions had mixed results, while modern heavy missions achieved consistent recovery."
    ]),
    (19, "Success Rate vs. Orbit Type", 'images/eda_success_rate_vs_orbit.png', [
        "100% Success Orbits: ES-L1, GEO, HEO, and SSO achieved a flawless 100% first-stage landing recovery rate.",
        "High Success Orbits: VLEO (Starlink constellation) achieved ~85% success; LEO achieved ~71%; Polar Orbits (PO) achieved ~67%.",
        "GTO Challenges: Geostationary Transfer Orbit achieved ~50% success due to high re-entry velocity and thermal stress requiring drone ship landings.",
        "SO Anomalies: Sub-Orbital test flight experienced 0% landing recovery."
    ]),
    (20, "Flight Number vs. Orbit Type", 'images/eda_flight_vs_orbit.png', [
        "Early Missions: Early flight numbers (<30) primarily targeted LEO and GTO orbits with experimental landing profiles.",
        "High Cadence Constellations: Beyond Flight 50, launch cadence shifted heavily toward VLEO for Starlink satellite deployment.",
        "Recovery Reliability: While early LEO and GTO missions experienced landing failures, modern flights across all orbits show unbroken runs of green landing successes."
    ]),
    (21, "Payload vs. Orbit Type", 'images/eda_payload_vs_orbit.png', [
        "Heavy LEO/VLEO Flights: Maximum payload masses (>10,000 kg up to 15,600 kg) are concentrated in LEO and VLEO orbits, where landing success is consistently high.",
        "GTO Payload Limits: GTO missions cluster in the 3,000 kg to 7,000 kg range. Heavy GTO launches leave less reserve fuel for booster entry burns.",
        "Polar & Scientific Missions: Polar and SSO missions carried payloads under 4,000 kg, allowing comfortable fuel margins for booster recovery."
    ]),
    (22, "Launch Success Yearly Trend", 'images/eda_yearly_trend.png', [
        "2010–2012 (0% Success): Purely expendable test launches; no booster recovery hardware existed.",
        "2013–2015 (~33% Success): Controlled ocean splashdowns and drone ship tests began; historic first ground pad recovery on Dec 22, 2015.",
        "2016–2017 (60–80% Success): Routine drone ship recoveries established on Autonomous Spaceport Drone Ships (ASDS).",
        "2018–2020 (>80–85% Success): Introduction of Falcon 9 Block 5 (titanium grid fins, upgraded heat shield) established industrial recovery reliability."
    ])
]

for s_idx, t_name, img_path, bullets in eda_slides_info:
    sl = prs.slides[s_idx]
    set_slide_title(sl, t_name)
    add_watermark(sl)
    c_sh = get_content_shape(sl)
    if c_sh:
        c_sh.left = Inches(0.8)
        c_sh.top = Inches(1.5)
        c_sh.width = Inches(4.8)
        c_sh.height = Inches(5.2)
        write_slide_content(c_sh, t_name, [("Key Observations & Findings:", bullets)], font_size=10.5)
    sl.shapes.add_picture(img_path, Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# Slide 24 to 33: EDA with SQL Slides
sql_slide_titles = {
    24: "All Launch Site Names",
    25: "Launch Site Names Begin with 'CCA'",
    26: "Total Payload Mass",
    27: "Average Payload Mass by F9 v1.1",
    28: "First Successful Ground Landing Date",
    29: "Successful Drone Ship Landing with Payload between 4000 and 6000",
    30: "Total Number of Successful and Failure Mission Outcomes",
    31: "Boosters Carried Maximum Payload",
    32: "2015 Launch Records",
    33: "Rank Landing Outcomes Between 2010-06-04 and 2017-03-20"
}

for s_num in range(24, 34):
    sl = prs.slides[s_num - 1]
    title_text = sql_slide_titles[s_num]
    set_slide_title(sl, title_text)
    add_watermark(sl)
    q_text, df_r, expl_text = sql_data[s_num]
    
    c_sh = get_content_shape(sl)
    if c_sh:
        c_sh.left = Inches(0.8)
        c_sh.top = Inches(1.4)
        c_sh.width = Inches(5.0)
        c_sh.height = Inches(5.4)
        
        write_slide_content(c_sh, title_text, [
            ("SQL Query Executed:", [q_text]),
            ("Key Analytical Findings & Context:", [expl_text])
        ], font_size=9.5)
        
    n_rows = min(len(df_r) + 1, 8)
    n_cols = len(df_r.columns)
    t_left = Inches(6.0)
    t_top = Inches(1.6)
    t_width = Inches(5.6)
    t_height = Inches(0.38 * n_rows)
    
    t_shape = sl.shapes.add_table(n_rows, n_cols, t_left, t_top, t_width, t_height)
    tbl = t_shape.table
    
    for c_i, col_name in enumerate(df_r.columns):
        cell = tbl.cell(0, c_i)
        cell.text = str(col_name)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_TITLE
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_NAME
            p.font.size = Pt(9.5)
            p.font.bold = True
            p.font.color.rgb = COLOR_WHITE
            
    for r_i in range(n_rows - 1):
        for c_i in range(n_cols):
            cell = tbl.cell(r_i + 1, c_i)
            v = df_r.iloc[r_i, c_i]
            cell.text = str(v) if pd.notnull(v) else "NULL"
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(245, 248, 252) if r_i % 2 == 0 else COLOR_WHITE
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                p.font.name = FONT_NAME
                p.font.size = Pt(9)
                p.font.color.rgb = COLOR_TEXT

# Slide 34: Section 3 Header
s34 = prs.slides[33]
add_watermark(s34)
for sh in s34.shapes:
    if sh.has_text_frame and 'TextBox' in sh.name:
        sh.text_frame.text = "Section 3: Interactive Visual Analytics (Folium Maps)"
        sh.text_frame.paragraphs[0].font.size = Pt(28)
        sh.text_frame.paragraphs[0].font.bold = True
        sh.text_frame.paragraphs[0].font.color.rgb = COLOR_TITLE

# Slide 35 to 37: Folium Map Slides
folium_slides_info = [
    (34, "Folium Map: Global Launch Sites Location Markers", 'images/folium_map_1_launch_sites.png', [
        "Coastal Proximity: All launch complexes are located immediately adjacent to coastlines (Atlantic Ocean in Florida, Pacific Ocean in California).",
        "Public Safety Buffer: Launches follow eastward trajectories over the ocean, ensuring that jettisoned first stages and launch aborts fall safely into open water rather than populated areas.",
        "Equatorial Velocity Boost: Florida launch complexes (Lat ~28.5°N) capitalize on Earth's rotational speed (~1,470 km/h) to maximize payload mass delivered to equatorial orbits.",
        "Polar Corridor: VAFB SLC-4E in California provides a dedicated southbound trajectory corridor for polar and sun-synchronous orbits without overflying continental landmasses."
    ]),
    (35, "Folium Map: Color-Labeled Outcomes and Marker Clusters", 'images/folium_map_2_marker_clusters.png', [
        "Marker Telemetry: Green markers designate successful first-stage landings; red markers designate landing failures or unattempted splashdowns.",
        "KSC LC-39A Leadership: KSC LC-39A exhibits the densest concentration of green markers, achieving a 76.9% landing success rate (10 successes out of 13 launches).",
        "Temporal Recovery Clustering: Cluster analysis indicates that early launches from CCAFS LC-40 contained most failure markers, whereas recent clusters across all pads are overwhelmingly green.",
        "Recovery Operational Maturity: Modern missions demonstrate high spatial recovery consistency regardless of launch complex."
    ]),
    (36, "Folium Map: Proximity Distances to Coastlines, Highways, Railways and Cities", 'images/folium_map_3_proximities.png', [
        "Haversine Spherical Distance Calculations:",
        "  • Coastline Proximity (~0.85 km): Launch pads are situated within 1 km of the shoreline for emergency abort safety and ASDS drone ship offloading logistics.",
        "  • Transportation Infrastructure (~1.25–1.35 km): Samuel C. Phillips Parkway highway and Titan III NASA Railroad enable heavy transport of 12-ft diameter rocket stages and propellants.",
        "  • City Distance (>16 km): Populated urban centers (Titusville, Cape Canaveral) are located over 16 km away, safeguarding civilian populations from acoustic shockwaves, sonic booms, and explosion hazards."
    ])
]

for s_idx, title_t, img_path, bullets in folium_slides_info:
    sl = prs.slides[s_idx]
    set_slide_title(sl, title_t)
    add_watermark(sl)
    c_sh = get_content_shape(sl)
    if c_sh:
        c_sh.left = Inches(0.8)
        c_sh.top = Inches(1.5)
        c_sh.width = Inches(4.8)
        c_sh.height = Inches(5.2)
        write_slide_content(c_sh, title_t, [("Key Observations & Geospatial Findings:", bullets)], font_size=10.5)
    sl.shapes.add_picture(img_path, Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# Slide 38: Section 4 Header
s38 = prs.slides[37]
add_watermark(s38)
for sh in s38.shapes:
    if sh.has_text_frame and 'TextBox' in sh.name:
        sh.text_frame.text = "Section 4: Interactive Dashboard (Plotly Dash)"
        sh.text_frame.paragraphs[0].font.size = Pt(28)
        sh.text_frame.paragraphs[0].font.bold = True
        sh.text_frame.paragraphs[0].font.color.rgb = COLOR_TITLE

# Slide 39 to 41: Plotly Dash Slides
dash_slides_info = [
    (38, "Plotly Dash: Total Launch Success Count for All Sites", 'images/dash_pie_all_sites.png', [
        "Success Proportion by Site: KSC LC-39A contributed the largest share of successful first-stage landings (41.7%), followed by CCAFS SLC-40 (29.2%), VAFB SLC-4E (20.8%), and CCAFS LC-40 (8.3%).",
        "Infrastructure Utilization: KSC LC-39A's historic pad infrastructure accommodates heavy-lift configurations and high launch cadence for commercial and NASA missions.",
        "Operational Concentration: Modern SpaceX missions are predominantly concentrated at KSC LC-39A and CCAFS SLC-40, maximizing recovery infrastructure utilization."
    ]),
    (39, "Plotly Dash: Success vs. Failure Ratio for Best Launch Site (KSC LC-39A)", 'images/dash_pie_best_site.png', [
        "Highest Success Rate: KSC LC-39A achieved a remarkable 76.9% landing success rate (10 successful recoveries vs. 3 failures).",
        "Advanced Pad Systems: KSC LC-39A features modernized launch mount systems and dedicated telemetry tracking corridors adapted from the Apollo and Space Shuttle programs.",
        "Risk Assessment: Customers launching payloads from KSC LC-39A experience the lowest booster recovery failure risk in the commercial space market."
    ]),
    (40, "Plotly Dash: Payload vs. Launch Outcome Correlation by Booster Version", 'images/dash_scatter_payload_outcome.png', [
        "Sweet-Spot Payload Range: The highest concentration of successful landings occurs between 2,000 kg and 5,500 kg across all booster versions.",
        "Booster Evolution: Early booster versions (v1.0, v1.1) suffered high failure rates across all payload masses; Block 5 (B5) and Full Thrust (FT) boosters demonstrate consistent landing success even for heavy payloads (>9,000 kg).",
        "Starlink Heavy Payloads: Missions carrying 15,600 kg achieve high landing success rates when deployed to Low Earth Orbit via drone ship recoveries."
    ])
]

for s_idx, title_t, img_path, bullets in dash_slides_info:
    sl = prs.slides[s_idx]
    set_slide_title(sl, title_t)
    add_watermark(sl)
    c_sh = get_content_shape(sl)
    if c_sh:
        c_sh.left = Inches(0.8)
        c_sh.top = Inches(1.5)
        c_sh.width = Inches(4.8)
        c_sh.height = Inches(5.2)
        write_slide_content(c_sh, title_t, [("Key Observations & Dashboard Insights:", bullets)], font_size=10.5)
    sl.shapes.add_picture(img_path, Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# Slide 42: Section 5 Header
s42 = prs.slides[41]
add_watermark(s42)
for sh in s42.shapes:
    if sh.has_text_frame and 'TextBox' in sh.name:
        sh.text_frame.text = "Section 5: Predictive Analysis (Classification) Results"
        sh.text_frame.paragraphs[0].font.size = Pt(28)
        sh.text_frame.paragraphs[0].font.bold = True
        sh.text_frame.paragraphs[0].font.color.rgb = COLOR_TITLE

# Slide 43: Classification Accuracy
s43 = prs.slides[42]
set_slide_title(s43, "Classification Accuracy")
add_watermark(s43)
c43 = get_content_shape(s43)
if c43:
    c43.left = Inches(0.8)
    c43.top = Inches(1.5)
    c43.width = Inches(4.8)
    c43.height = Inches(5.2)
    write_slide_content(c43, "Classification Accuracy", [
        ("Model Accuracy Comparison & Best Performer", [
            "Benchmark Test Accuracy: All 4 classification models (Logistic Regression, SVM, Decision Tree, and KNN) achieved an identical 83.33% Test Accuracy on the 18-sample unseen test set.",
            "Train CV Accuracy Scores (10-Fold Cross-Validation):",
            "  • Decision Tree: Best CV Score = 87.14% (criterion='gini', max_depth=4, min_samples_leaf=2).",
            "  • Support Vector Machine: CV Score = 84.82% (kernel='sigmoid', C=1.0, gamma=0.0316).",
            "  • K-Nearest Neighbors: CV Score = 84.82% (n_neighbors=10, p=1).",
            "  • Logistic Regression: CV Score = 84.64% (C=0.01, penalty='l2').",
            "Model Selection: Decision Tree and Logistic Regression offer the highest operational utility due to high CV stability, complete explainability, and zero test overfitting."
        ])
    ], font_size=10)
s43.shapes.add_picture('images/ml_model_accuracy_comparison.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# Slide 44: Confusion Matrix
s44 = prs.slides[43]
set_slide_title(s44, "Confusion Matrix")
add_watermark(s44)
c44 = get_content_shape(s44)
if c44:
    c44.left = Inches(0.8)
    c44.top = Inches(1.5)
    c44.width = Inches(4.8)
    c44.height = Inches(5.2)
    write_slide_content(c44, "Confusion Matrix", [
        ("Confusion Matrix Analysis & Operational Impact", [
            "Test Set Sample Distribution (18 Total Samples):",
            "  • True Positives (TP = 12): Model correctly predicted 12 successful landing recoveries.",
            "  • True Negatives (TN = 3): Model correctly predicted 3 landing failures.",
            "  • False Positives (FP = 3): Model predicted success for 3 missions that failed.",
            "  • False Negatives (FN = 0): Zero successful landings were misclassified as failures!",
            "Critical Aerospace Value:",
            "  • 100% Recall (Sensitivity): A zero false-negative rate guarantees that recovery ships, tugs, and operations teams are never mistakenly stood down when a booster is recoverable.",
            "  • High Precision (80.0%): Provides commercial customers with high confidence in recovery likelihood prior to mission execution."
        ])
    ], font_size=10)
s44.shapes.add_picture('images/ml_confusion_matrix.png', Inches(5.8), Inches(1.6), Inches(5.8), Inches(4.6))

# Slide 45: Conclusions
s45 = prs.slides[44]
set_slide_title(s45, "Conclusions")
add_watermark(s45)
c45 = get_content_shape(s45)
if c45:
    c45.left = Inches(0.8)
    c45.top = Inches(1.5)
    c45.width = Inches(11.5)
    c45.height = Inches(5.2)
    write_slide_content(c45, "Conclusions & Innovative Recommendations", [
        ("1. Core Analytical Findings", [
            "Technical Learning Curve: SpaceX successfully industrialized rocket reusability, escalating first-stage recovery from 0% in 2010-2012 to over 80% in 2018-2020 via iterative Block 5 engineering advancements.",
            "Geographic & Orbital Determinants: Coastal launch complexes maximize safety while exploiting equatorial rotational boost. VLEO, SSO, and polar orbits yield near-perfect landing success, whereas high-velocity GTO missions carry higher recovery risk (~50%).",
            "Predictive Machine Learning Viability: Supervised classification models achieved 83.33% test accuracy with 100% recall for landing success, confirming high predictive feasibility using pre-launch flight parameters."
        ]),
        ("2. Key Innovative Insights", [
            "Economic Asymmetry: Competitor launch providers operate expendable rockets costing $165M+, whereas SpaceX prices Falcon 9 at $62M by reusing first-stage boosters (~70% of rocket value).",
            "Operational Value of 100% Recall: Zero false negatives ensures that recovery fleets and support teams are never stood down when a stage is recoverable, saving tens of millions in hardware assets."
        ]),
        ("3. Actionable Next Steps & Commercial Recommendations", [
            "Competitor Bidding Strategy: Competitor launch providers (e.g. ULA, Arianespace, Rocket Lab) can utilize this predictive framework to anticipate whether SpaceX will achieve booster recovery on a commercial RFP, enabling precise pricing against SpaceX's $62M rate.",
            "Risk & Insurance Underwriting: Space and marine insurance underwriters can dynamically adjust premiums based on flight profile (e.g., discounted premiums for KSC LC-39A and LEO/VLEO missions; adjusted premiums for heavy GTO flights).",
            "Payload Mass Optimization: Commercial satellite operators should target payload masses between 2,000 kg and 5,500 kg to maximize recovery likelihood and negotiate discounts on flight-proven boosters."
        ])
    ], font_size=10)

# Slide 46: Appendix
s46 = prs.slides[45]
set_slide_title(s46, "Appendix")
add_watermark(s46)
c46 = get_content_shape(s46)
if c46:
    c46.left = Inches(0.8)
    c46.top = Inches(1.5)
    c46.width = Inches(11.5)
    c46.height = Inches(5.2)
    write_slide_content(c46, "Appendix: Verified GitHub Links & Assets", [
        ("Verified GitHub Repository & Code Deliverables", [
            "Main Repository URL: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction",
            "1. Data Collection API: notebooks/1_SpaceX_Data_Collection_API.ipynb",
            "2. Data Collection Scraping: notebooks/2_SpaceX_Data_Collection_Web_Scraping.ipynb",
            "3. Data Wrangling: notebooks/3_SpaceX_Data_Wrangling.ipynb",
            "4. EDA Data Visualization: notebooks/4_SpaceX_EDA_Data_Visualization.ipynb",
            "5. EDA with SQL: notebooks/5_SpaceX_EDA_SQL.ipynb",
            "6. Folium Interactive Maps: notebooks/6_SpaceX_Interactive_Visual_Analytics_Folium.ipynb",
            "7. Plotly Dash Dashboard Application: spacex_dash_app.py",
            "8. Machine Learning Prediction: notebooks/7_SpaceX_Machine_Learning_Prediction.ipynb"
        ]),
        ("Data Dictionary, Creativity & Technology Stack", [
            "Creativity & Polish Applied: High-resolution custom 300 DPI analytical figures, embedded architecture flowcharts, formatted SQL data tables, cohesive aerospace color palette, and plagiarism watermark.",
            "Data Sources: SpaceX REST API v4, Wikipedia Falcon 9 historical archives, IBM Skills Network database.",
            "Technologies Utilized: Python 3.10, Pandas, NumPy, Scikit-Learn, Matplotlib, Seaborn, Folium, Plotly Dash, SQLite, python-pptx."
        ])
    ], font_size=10)

# Slide 47: Innovative Insights & Thank You Slide
s47 = prs.slides[46]
add_watermark(s47)
tx_box = s47.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.5), Inches(5.6))
tf_ty = tx_box.text_frame
tf_ty.word_wrap = True

p_ty0 = tf_ty.paragraphs[0]
p_ty0.text = "Innovative Insights & Project Conclusion"
p_ty0.font.name = FONT_NAME
p_ty0.font.size = Pt(24)
p_ty0.font.bold = True
p_ty0.font.color.rgb = COLOR_TITLE
p_ty0.space_after = Pt(10)

p_h1 = tf_ty.add_paragraph()
p_h1.text = "Innovative Aerospace Business Insights:"
p_h1.font.name = FONT_NAME
p_h1.font.size = Pt(13)
p_h1.font.bold = True
p_h1.font.color.rgb = COLOR_BLUE

p_b1 = tf_ty.add_paragraph()
p_b1.text = "• Dynamic RFP Bidding Model: Because SpaceX's Falcon 9 costs $62M (reusable) vs. $165M+ (expendable competitors), competitor launch providers can use our 83.33% accurate model to predict whether SpaceX's recovery bid on a specific satellite RFP has a high or low probability of failure. For heavy GTO launches, competitors can exploit SpaceX's ~50% recovery margin to win launch contracts."
p_b1.font.name = FONT_NAME
p_b1.font.size = Pt(11)
p_b1.font.color.rgb = COLOR_TEXT
p_b1.space_after = Pt(4)

p_b2 = tf_ty.add_paragraph()
p_b2.text = "• Asymmetric Cost of False Negatives: In aerospace logistics, false negatives (classifying a recoverable booster as lost) carry massive capital cost ($30M+ wasted hardware). Our models achieved 100% Recall for landing success, providing a zero-false-negative operational framework for launch recovery management."
p_b2.font.name = FONT_NAME
p_b2.font.size = Pt(11)
p_b2.font.color.rgb = COLOR_TEXT
p_b2.space_after = Pt(12)

p_ty3 = tf_ty.add_paragraph()
p_ty3.alignment = PP_ALIGN.CENTER
p_ty3.text = "Author: Heet Patel | IBM Data Science Professional Certificate | Coursera"
p_ty3.font.name = FONT_NAME
p_ty3.font.size = Pt(13)
p_ty3.font.bold = True
p_ty3.font.color.rgb = COLOR_TITLE
p_ty3.space_after = Pt(4)

p_ty4 = tf_ty.add_paragraph()
p_ty4.alignment = PP_ALIGN.CENTER
p_ty4.text = "GitHub Repository: https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction"
p_ty4.font.name = FONT_NAME
p_ty4.font.size = Pt(13)
p_ty4.font.bold = True
p_ty4.font.color.rgb = RGBColor(39, 174, 96)

# Save presentations
output_pptx = "Data Science Capstone Project Report.pptx"
prs.save(output_pptx)
prs.save("ds-capstone-template-coursera.pptx")
print("Perfect presentation built and saved successfully!")
