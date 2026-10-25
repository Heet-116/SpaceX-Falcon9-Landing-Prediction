# 🚀 SpaceX Falcon 9 First Stage Landing Prediction
### IBM Applied Data Science Capstone Project
**Author:** Heet Patel ([@Heet-116](https://github.com/Heet-116))  
**Repository:** [https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction](https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction)

---

## 📌 Executive Summary
SpaceX has revolutionized the commercial aerospace industry by dramatically lowering launch costs. A standard Falcon 9 launch costs **$62 million**, compared to **$165+ million** from other providers, primarily because SpaceX recovers and reuses the Falcon 9 first stage booster. 

In this capstone project, we apply end-to-end data science methodologies to predict whether a Falcon 9 first stage will successfully land. This enables commercial competitors and satellite operators to forecast SpaceX launch pricing, determine first stage recovery probabilities, and formulate competitive bidding strategies.

Key Findings:
- **Overall Landing Success Rate:** 66.7% across 90 analyzed Falcon 9 missions.
- **Success Trajectory:** Escalated from 0% (2010–2012) to **>80%** (2018–2020) with Block 5 booster reusability.
- **Top Launch Complex:** **KSC LC-39A** achieved the highest success rate of **76.9%**.
- **Best Machine Learning Models:** All four classification models (Logistic Regression, Support Vector Machine, Decision Tree, and K-Nearest Neighbors) achieved **83.33% Test Accuracy** with **100% Recall** for landing successes.

---

## 🏗️ Project Architecture & Pipeline
```
SpaceX REST API  ----+
                     +--> Data Wrangling & Feature Engineering --> SQLite / EDA
Wikipedia Scraping --+        (Handling Nulls, One-Hot Encoding)
                                             |
                                             v
                             Interactive Visual Analytics
                               (Folium Maps & Plotly Dash)
                                             |
                                             v
                             Supervised Machine Learning
                          (GridSearchCV Hyperparameter Tuning)
                                             |
                                             v
                           Executive Presentation & PDF Report
```

---

## 📂 Repository Contents
| File / Directory | Description |
| :--- | :--- |
| [`notebooks/1_SpaceX_Data_Collection_API.ipynb`](notebooks/1_SpaceX_Data_Collection_API.ipynb) | SpaceX REST API queries and JSON response parsing. |
| [`notebooks/2_SpaceX_Data_Collection_Web_Scraping.ipynb`](notebooks/2_SpaceX_Data_Collection_Web_Scraping.ipynb) | BeautifulSoup web scraping of Falcon 9 launch records from Wikipedia. |
| [`notebooks/3_SpaceX_Data_Wrangling.ipynb`](notebooks/3_SpaceX_Data_Wrangling.ipynb) | Handling missing values, landing outcome categorization, target encoding. |
| [`notebooks/4_SpaceX_EDA_Data_Visualization.ipynb`](notebooks/4_SpaceX_EDA_Data_Visualization.ipynb) | Exploratory data analysis using Pandas, Matplotlib, and Seaborn. |
| [`notebooks/5_SpaceX_EDA_SQL.ipynb`](notebooks/5_SpaceX_EDA_SQL.ipynb) | Exploratory data analysis via SQL queries on SQLite database. |
| [`notebooks/6_SpaceX_Interactive_Visual_Analytics_Folium.ipynb`](notebooks/6_SpaceX_Interactive_Visual_Analytics_Folium.ipynb) | Geospatial launch site mapping, marker clustering, proximity analysis. |
| [`spacex_dash_app.py`](spacex_dash_app.py) | Interactive dashboard using Plotly Dash with callbacks and sliders. |
| [`notebooks/7_SpaceX_Machine_Learning_Prediction.ipynb`](notebooks/7_SpaceX_Machine_Learning_Prediction.ipynb) | Machine learning classification models with GridSearchCV tuning. |
| `data/` | Historical SpaceX launch datasets (Part 1, 2, 3, Dash, Geo, SQL). |
| `images/` | High-resolution 300 DPI analytical charts, maps, and flowcharts. |
| `Data Science Capstone Project Report.pdf` | Final 47-slide technical presentation submitted in PDF format. |

---

## 🧪 Machine Learning Model Comparison
| Algorithm | Best Hyperparameters | Train CV Accuracy | Test Set Accuracy |
| :--- | :--- | :---: | :---: |
| **Logistic Regression** | `{'C': 0.01, 'penalty': 'l2', 'solver': 'lbfgs'}` | **84.64%** | **83.33%** |
| **Support Vector Machine (SVM)** | `{'C': 1.0, 'gamma': 0.0316, 'kernel': 'sigmoid'}` | **84.82%** | **83.33%** |
| **Decision Tree Classifier** | `{'criterion': 'gini', 'max_depth': 4, 'min_samples_leaf': 2}` | **87.14%** | **83.33%** |
| **K-Nearest Neighbors (KNN)** | `{'algorithm': 'auto', 'n_neighbors': 10, 'p': 1}` | **84.82%** | **83.33%** |

### Confusion Matrix Highlights:
- **True Positives:** 12
- **True Negatives:** 3
- **False Positives:** 3
- **False Negatives:** 0
- **Recall (Success):** 100%

---

## 📊 How to Run the Dash Application
```bash
# Clone the repository
git clone https://github.com/Heet-116/SpaceX-Falcon9-Landing-Prediction.git
cd SpaceX-Falcon9-Landing-Prediction

# Install dependencies
pip install pandas dash plotly

# Run the Plotly Dash application
python spacex_dash_app.py
```
Open your browser and navigate to `http://127.0.0.1:8050/`.

---

## 📜 License & Acknowledgements
This project is completed as part of the **IBM Data Science Professional Certificate** on Coursera.  
Data provided by SpaceX and IBM Skills Network.
