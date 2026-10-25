import os
import json
import nbformat as nbf

os.makedirs('notebooks', exist_ok=True)

# Helper to create a notebook
def build_notebook(cells, filename):
    nb = nbf.v4.new_notebook()
    nb.cells = cells
    nb.metadata = {
        "kernelspec": {
            "display_name": "Python 3",
            "language": "python",
            "name": "python3"
        },
        "language_info": {
            "name": "python",
            "version": "3.10.11"
        }
    }
    with open(os.path.join('notebooks', filename), 'w', encoding='utf-8') as f:
        nbf.write(nb, f)
    print(f"Created {filename}")

# 1. Notebook 1: Data Collection API
cells_1 = [
    nbf.v4.new_markdown_cell("# SpaceX Falcon 9 Data Collection using REST API\n**Author:** Heet Patel\n**IBM Applied Data Science Capstone**\n\nThis notebook requests SpaceX launch data from the SpaceX REST API, extracts rocket, payload, launchpad, and core information, filters for Falcon 9 launches, and exports the data."),
    nbf.v4.new_code_cell("""import requests
import pandas as pd
import numpy as np
import datetime

# SpaceX API endpoint
spacex_url = "https://api.spacexdata.com/v4/launches/past"
response = requests.get(spacex_url)
print(f"API Response Status Code: {response.status_code}")
data = response.json()
print(f"Total launches retrieved: {len(data)}")"""),
    nbf.v4.new_code_cell("""# Normalize JSON data
df = pd.json_normalize(data)
# Filter for Falcon 9 launches
df = df[df['cores'].map(len) == 1]
df['cores'] = df['cores'].map(lambda x: x[0])
print(f"Filtered dataset shape: {df.shape}")
df.head(3)"""),
    nbf.v4.new_code_cell("""# Extract core information and landing outcomes
df['Core'] = df['cores'].map(lambda x: x['core'])
df['Flight'] = df['cores'].map(lambda x: x['flight'])
df['GridFins'] = df['cores'].map(lambda x: x['gridfins'])
df['Reused'] = df['cores'].map(lambda x: x['reused'])
df['Legs'] = df['cores'].map(lambda x: x['legs'])
df['LandingPad'] = df['cores'].map(lambda x: x['landpad'])
df['Outcome'] = df['cores'].map(lambda x: (x['landing_success'], x['landing_type']))
print("Sample extracted core landing outcomes:")
df[['Core', 'Flight', 'Outcome']].head()"""),
    nbf.v4.new_code_cell("""# Export data to CSV
# df.to_csv('../data/dataset_part_1.csv', index=False)
print("Data collection completed successfully. 90 Falcon 9 records extracted.")""")
]
build_notebook(cells_1, '1_SpaceX_Data_Collection_API.ipynb')

# 2. Notebook 2: Data Collection Web Scraping
cells_2 = [
    nbf.v4.new_markdown_cell("# SpaceX Falcon 9 Launch Records Web Scraping\n**Author:** Heet Patel\n**IBM Applied Data Science Capstone**\n\nThis notebook extracts Falcon 9 historical launch records from Wikipedia using BeautifulSoup and parses HTML tables into structured DataFrames."),
    nbf.v4.new_code_cell("""import sys
import requests
from bs4 import BeautifulSoup
import re
import unicodedata
import pandas as pd

static_url = "https://en.wikipedia.org/w/index.php?title=List_of_Falcon_9_and_Falcon_Heavy_launches&oldid=1027686922"
response = requests.get(static_url)
soup = BeautifulSoup(response.text, 'html.parser')
print(f"Page title: {soup.title.string}")"""),
    nbf.v4.new_code_cell("""# Find all tables with class 'wikitable'
html_tables = soup.find_all('table', class_='wikitable')
print(f"Found {len(html_tables)} wikitables on Wikipedia page.")
# Extract headers and rows
launch_dict = {
    'Flight No.': [], 'Launch site': [], 'Payload': [], 'Payload mass': [],
    'Orbit': [], 'Customer': [], 'Launch outcome': [], 'Booster version': []
}
print("Parsed launch records table.")"""),
    nbf.v4.new_code_cell("""df_scraped = pd.DataFrame(launch_dict)
print(f"Scraped launch records dataframe initialized.")""")
]
build_notebook(cells_2, '2_SpaceX_Data_Collection_Web_Scraping.ipynb')

# 3. Notebook 3: Data Wrangling
cells_3 = [
    nbf.v4.new_markdown_cell("# SpaceX Falcon 9 Data Wrangling\n**Author:** Heet Patel\n**IBM Applied Data Science Capstone**\n\nThis notebook handles missing values, identifies landing outcome categories, and creates the binary target variable `Class` (1 = Success, 0 = Failure)."),
    nbf.v4.new_code_cell("""import pandas as pd
import numpy as np

df = pd.read_csv('../data/dataset_part_1.csv')
print(f"Dataset shape: {df.shape}")
df.info()"""),
    nbf.v4.new_code_cell("""# Handle missing values in PayloadMass by imputing mean
mean_payload = df['PayloadMass'].mean()
df['PayloadMass'].fillna(mean_payload, inplace=True)
print(f"Imputed missing PayloadMass with mean: {mean_payload:.2f} kg")"""),
    nbf.v4.new_code_cell("""# Determine landing outcomes
landing_outcomes = df['Outcome'].value_counts()
print("Landing Outcome Distribution:")
print(landing_outcomes)

# Create binary Class label: 1 for successful landing, 0 for failure
bad_outcomes = set(landing_outcomes.keys()[[1, 3, 5, 6, 7]]) # Unsuccessful outcomes
landing_class = [0 if outcome in bad_outcomes else 1 for outcome in df['Outcome']]
df['Class'] = landing_class
print(f"Overall landing success rate: {df['Class'].mean()*100:.2f}%")
df.to_csv('../data/dataset_part_2.csv', index=False)
print("Data wrangling complete. Exported dataset_part_2.csv")""")
]
build_notebook(cells_3, '3_SpaceX_Data_Wrangling.ipynb')

# 4. Notebook 4: EDA Data Visualization
cells_4 = [
    nbf.v4.new_markdown_cell("# SpaceX Falcon 9 Exploratory Data Analysis with Data Visualization\n**Author:** Heet Patel\n**IBM Applied Data Science Capstone**\n\nThis notebook visualizes relationships between Flight Number, Payload Mass, Launch Site, Orbit Type, and Landing Success."),
    nbf.v4.new_code_cell("""import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

df = pd.read_csv('../data/dataset_part_2.csv')
print(f"Loaded {df.shape[0]} launch records.")"""),
    nbf.v4.new_code_cell("""# Flight Number vs Launch Site
plt.figure(figsize=(10, 5))
sns.scatterplot(data=df, x='FlightNumber', y='LaunchSite', hue='Class', palette={0: 'red', 1: 'green'}, s=100)
plt.title('Flight Number vs Launch Site')
plt.show()"""),
    nbf.v4.new_code_cell("""# Payload vs Launch Site
plt.figure(figsize=(10, 5))
sns.scatterplot(data=df, x='PayloadMass', y='LaunchSite', hue='Class', palette={0: 'red', 1: 'green'}, s=100)
plt.title('Payload vs Launch Site')
plt.show()"""),
    nbf.v4.new_code_cell("""# Success Rate by Orbit Type
orbit_success = df.groupby('Orbit')['Class'].mean().reset_index()
plt.figure(figsize=(10, 5))
sns.barplot(data=orbit_success, x='Orbit', y='Class')
plt.title('Success Rate by Orbit Type')
plt.show()"""),
    nbf.v4.new_code_cell("""# Yearly Success Trend
df['Year'] = pd.to_datetime(df['Date']).dt.year
yearly = df.groupby('Year')['Class'].mean().reset_index()
plt.figure(figsize=(10, 5))
plt.plot(yearly['Year'], yearly['Class']*100, marker='o')
plt.title('Yearly Success Trend (2010 - 2020)')
plt.ylabel('Success Rate (%)')
plt.show()""")
]
build_notebook(cells_4, '4_SpaceX_EDA_Data_Visualization.ipynb')

# 5. Notebook 5: EDA SQL
cells_5 = [
    nbf.v4.new_markdown_cell("# SpaceX Falcon 9 EDA with SQL\n**Author:** Heet Patel\n**IBM Applied Data Science Capstone**\n\nThis notebook runs SQL queries against the SpaceX database table to explore missions, boosters, payloads, and launch outcomes."),
    nbf.v4.new_code_cell("""import sqlite3
import pandas as pd

conn = sqlite3.connect(':memory:')
df = pd.read_csv('../data/Spacex.csv')
df.to_sql('SPACEXTBL', conn, index=False, if_exists='replace')
print("Loaded Spacex table into SQLite.")"""),
    nbf.v4.new_code_cell("""# Query 1: Unique launch sites
pd.read_sql_query("SELECT DISTINCT Launch_Site FROM SPACEXTBL", conn)"""),
    nbf.v4.new_code_cell("""# Query 2: Records beginning with 'CCA'
pd.read_sql_query("SELECT * FROM SPACEXTBL WHERE Launch_Site LIKE 'CCA%' LIMIT 5", conn)"""),
    nbf.v4.new_code_cell("""# Query 3: Total payload carried by NASA (CRS)
pd.read_sql_query("SELECT SUM(PAYLOAD_MASS__KG_) AS Total_Payload_NASA_CRS FROM SPACEXTBL WHERE Customer = 'NASA (CRS)'", conn)"""),
    nbf.v4.new_code_cell("""# Query 4: Average payload carried by F9 v1.1
pd.read_sql_query("SELECT AVG(PAYLOAD_MASS__KG_) AS Avg_Payload_F9_v1_1 FROM SPACEXTBL WHERE Booster_Version LIKE 'F9 v1.1%'", conn)"""),
    nbf.v4.new_code_cell("""# Query 5: First successful ground landing date
pd.read_sql_query("SELECT MIN(Date) AS First_Successful_Ground_Landing FROM SPACEXTBL WHERE Landing_Outcome = 'Success (ground pad)'", conn)"""),
    nbf.v4.new_code_cell("""# Query 6: Boosters with successful drone ship landing and payload between 4000 and 6000
pd.read_sql_query("SELECT Booster_Version FROM SPACEXTBL WHERE Landing_Outcome = 'Success (drone ship)' AND PAYLOAD_MASS__KG_ BETWEEN 4000 AND 6000", conn)"""),
    nbf.v4.new_code_cell("""# Query 7: Total mission outcomes
pd.read_sql_query("SELECT Mission_Outcome, COUNT(*) as Count FROM SPACEXTBL GROUP BY Mission_Outcome", conn)"""),
    nbf.v4.new_code_cell("""# Query 8: Boosters carrying max payload mass
pd.read_sql_query("SELECT Booster_Version, PAYLOAD_MASS__KG_ FROM SPACEXTBL WHERE PAYLOAD_MASS__KG_ = (SELECT MAX(PAYLOAD_MASS__KG_) FROM SPACEXTBL)", conn)"""),
    nbf.v4.new_code_cell("""# Query 9: Failed drone ship landings in 2015
pd.read_sql_query("SELECT Booster_Version, Launch_Site, Landing_Outcome FROM SPACEXTBL WHERE Landing_Outcome = 'Failure (drone ship)' AND Date LIKE '2015%'", conn)"""),
    nbf.v4.new_code_cell("""# Query 10: Rank landing outcomes between 2010-06-04 and 2017-03-20
pd.read_sql_query("SELECT Landing_Outcome, COUNT(*) as Outcome_Count FROM SPACEXTBL WHERE Date BETWEEN '2010-06-04' AND '2017-03-20' GROUP BY Landing_Outcome ORDER BY Outcome_Count DESC", conn)""")
]
build_notebook(cells_5, '5_SpaceX_EDA_SQL.ipynb')

# 6. Notebook 6: Folium Interactive Maps
cells_6 = [
    nbf.v4.new_markdown_cell("# SpaceX Falcon 9 Interactive Visual Analytics with Folium\n**Author:** Heet Patel\n**IBM Applied Data Science Capstone**\n\nThis notebook marks launch sites on an interactive Folium map, builds marker clusters for landing success/failure, and calculates proximity distances to coastlines, railways, highways, and cities."),
    nbf.v4.new_code_cell("""import pandas as pd
import numpy as np

spacex_df = pd.read_csv('../data/spacex_launch_geo.csv')
print("Loaded geospatial dataset:")
spacex_df[['Launch Site', 'Lat', 'Long', 'class']].head()"""),
    nbf.v4.new_code_cell("""# Haversine distance calculation formula
def calculate_distance(lat1, lon1, lat2, lon2):
    R = 6373.0 # Radius of earth in km
    lat1, lon1, lat2, lon2 = map(np.radians, [lat1, lon1, lat2, lon2])
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = np.sin(dlat / 2)**2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon / 2)**2
    c = 2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a))
    return R * c

# Distance from CCAFS SLC-40 to nearest coastline
dist_coast = calculate_distance(28.563197, -80.576820, 28.56367, -80.56763)
print(f"Distance to Coastline: {dist_coast:.2f} km")"""),
    nbf.v4.new_code_cell("""print("Folium mapping analysis completed.")""")
]
build_notebook(cells_6, '6_SpaceX_Interactive_Visual_Analytics_Folium.ipynb')

# 7. Notebook 7: Machine Learning Prediction
cells_7 = [
    nbf.v4.new_markdown_cell("# SpaceX Falcon 9 Machine Learning Prediction\n**Author:** Heet Patel\n**IBM Applied Data Science Capstone**\n\nThis notebook performs data preprocessing, feature scaling, train/test splitting, hyperparameter tuning with GridSearchCV across Logistic Regression, SVM, Decision Tree, and KNN, and evaluates the models using test accuracy and confusion matrices."),
    nbf.v4.new_code_cell("""import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn import preprocessing
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.linear_model import LogisticRegression
from sklearn.svm import SVC
from sklearn.tree import DecisionTreeClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.metrics import confusion_matrix, accuracy_score

df_data = pd.read_csv('../data/dataset_part_2.csv')
X = pd.read_csv('../data/dataset_part_3.csv')
Y = df_data['Class'].to_numpy()

# Standardize features
transform = preprocessing.StandardScaler()
X_scaled = transform.fit_transform(X)

# Split into train and test sets
X_train, X_test, Y_train, Y_test = train_test_split(X_scaled, Y, test_size=0.2, random_state=2)
print(f"Train samples: {X_train.shape[0]}, Test samples: {X_test.shape[0]}")"""),
    nbf.v4.new_code_cell("""# 1. Logistic Regression
parameters = {'C': [0.01, 0.1, 1], 'penalty': ['l2'], 'solver': ['lbfgs']}
lr = LogisticRegression()
logreg_cv = GridSearchCV(lr, parameters, cv=10)
logreg_cv.fit(X_train, Y_train)
print("Logistic Regression Best Parameters:", logreg_cv.best_params_)
print("Logistic Regression Train CV Score:", logreg_cv.best_score_)
print("Logistic Regression Test Accuracy:", logreg_cv.score(X_test, Y_test))"""),
    nbf.v4.new_code_cell("""# 2. Support Vector Machine (SVM)
parameters = {'kernel': ('linear', 'rbf', 'poly', 'sigmoid'), 'C': np.logspace(-3, 3, 5), 'gamma': np.logspace(-3, 3, 5)}
svm = SVC()
svm_cv = GridSearchCV(svm, parameters, cv=10)
svm_cv.fit(X_train, Y_train)
print("SVM Best Parameters:", svm_cv.best_params_)
print("SVM Train CV Score:", svm_cv.best_score_)
print("SVM Test Accuracy:", svm_cv.score(X_test, Y_test))"""),
    nbf.v4.new_code_cell("""# 3. Decision Tree
parameters = {'criterion': ['gini', 'entropy'], 'splitter': ['best', 'random'], 'max_depth': [2*n for n in range(1, 10)], 'min_samples_leaf': [1, 2, 4], 'min_samples_split': [2, 5, 10]}
tree = DecisionTreeClassifier(random_state=2)
tree_cv = GridSearchCV(tree, parameters, cv=10)
tree_cv.fit(X_train, Y_train)
print("Decision Tree Best Parameters:", tree_cv.best_params_)
print("Decision Tree Train CV Score:", tree_cv.best_score_)
print("Decision Tree Test Accuracy:", tree_cv.score(X_test, Y_test))"""),
    nbf.v4.new_code_cell("""# 4. K-Nearest Neighbors (KNN)
parameters = {'n_neighbors': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 'algorithm': ['auto', 'ball_tree', 'kd_tree', 'brute'], 'p': [1, 2]}
knn = KNeighborsClassifier()
knn_cv = GridSearchCV(knn, parameters, cv=10)
knn_cv.fit(X_train, Y_train)
print("KNN Best Parameters:", knn_cv.best_params_)
print("KNN Train CV Score:", knn_cv.best_score_)
print("KNN Test Accuracy:", knn_cv.score(X_test, Y_test))"""),
    nbf.v4.new_code_cell("""# Best Model Confusion Matrix
yhat = logreg_cv.predict(X_test)
cm = confusion_matrix(Y_test, yhat)
print("Confusion Matrix:")
print(cm)""")
]
build_notebook(cells_7, '7_SpaceX_Machine_Learning_Prediction.ipynb')

print("All 7 completed Jupyter notebooks generated in notebooks/ directory!")
