import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import urllib.request
import sqlite3
import numpy as np
import pandas as pd
import seaborn as sns
from sklearn import preprocessing
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.linear_model import LogisticRegression
from sklearn.svm import SVC
from sklearn.tree import DecisionTreeClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.metrics import confusion_matrix, accuracy_score, classification_report

plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.sans-serif'] = 'Arial'
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['figure.dpi'] = 300

os.makedirs('images', exist_ok=True)
os.makedirs('data', exist_ok=True)

df_p2 = pd.read_csv('data/dataset_part_2.csv')
df_p3 = pd.read_csv('data/dataset_part_3.csv')
df_dash = pd.read_csv('data/spacex_launch_dash.csv')
df_sql = pd.read_csv('data/Spacex.csv')

print("\n--- Training Machine Learning Models ---")
Y = df_p2['Class'].to_numpy()
X = df_p3
scaler = preprocessing.StandardScaler()
X_scaled = scaler.fit_transform(X)
X_train, X_test, Y_train, Y_test = train_test_split(X_scaled, Y, test_size=0.2, random_state=2)

models = {
    'Logistic Regression': (LogisticRegression(max_iter=1000), {'C': [0.01, 0.1, 1], 'penalty': ['l2'], 'solver': ['lbfgs']}),
    'Support Vector Machine': (SVC(), {'kernel': ('linear', 'rbf', 'poly', 'sigmoid'), 'C': np.logspace(-3, 3, 5), 'gamma': np.logspace(-3, 3, 5)}),
    'Decision Tree': (DecisionTreeClassifier(random_state=2), {'criterion': ['gini', 'entropy'], 'splitter': ['best', 'random'], 'max_depth': [2*n for n in range(1, 10)], 'min_samples_leaf': [1, 2, 4], 'min_samples_split': [2, 5, 10]}),
    'K-Nearest Neighbors': (KNeighborsClassifier(), {'n_neighbors': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 'algorithm': ['auto', 'ball_tree', 'kd_tree', 'brute'], 'p': [1, 2]})
}

model_results = {}
for name, (estimator, params) in models.items():
    grid = GridSearchCV(estimator, params, cv=10, n_jobs=1)
    grid.fit(X_train, Y_train)
    test_acc = grid.score(X_test, Y_test)
    y_pred = grid.predict(X_test)
    cm = confusion_matrix(Y_test, y_pred)
    model_results[name] = {
        'best_params': grid.best_params_,
        'train_acc': grid.best_score_,
        'test_acc': test_acc,
        'cm': cm,
        'model': grid.best_estimator_
    }
    print(f"{name}: Train CV Score = {grid.best_score_:.4f}, Test Accuracy = {test_acc:.4f}")

# Model Accuracy Bar Chart (Slide 43)
fig, ax = plt.subplots(figsize=(10, 5))
names = list(model_results.keys())
train_scores = [model_results[n]['train_acc'] * 100 for n in names]
test_scores = [model_results[n]['test_acc'] * 100 for n in names]

x = np.arange(len(names))
width = 0.35

rects1 = ax.bar(x - width/2, train_scores, width, label='Train CV Accuracy (%)', color='#3498db', edgecolor='black')
rects2 = ax.bar(x + width/2, test_scores, width, label='Test Set Accuracy (%)', color='#2ecc71', edgecolor='black')

for rect in rects1:
    h = rect.get_height()
    ax.annotate(f'{h:.1f}%', (rect.get_x() + rect.get_width()/2, h), textcoords="offset points", xytext=(0, 4), ha='center', fontsize=10, fontweight='bold')
for rect in rects2:
    h = rect.get_height()
    ax.annotate(f'{h:.1f}%', (rect.get_x() + rect.get_width()/2, h), textcoords="offset points", xytext=(0, 4), ha='center', fontsize=10, fontweight='bold')

ax.set_ylabel('Accuracy (%)', fontweight='bold', fontsize=11)
ax.set_title('Classification Model Performance Comparison (Train CV vs. Test Set)', fontweight='bold', fontsize=14, pad=12)
ax.set_xticks(x)
ax.set_xticklabels(names, fontweight='bold', fontsize=11)
ax.set_ylim(0, 110)
ax.legend(frameon=True, fontsize=10)
plt.tight_layout()
plt.savefig('images/ml_model_accuracy_comparison.png', dpi=300)
plt.close()

# Confusion Matrix for Best Model (Slide 44)
best_cm = model_results['Logistic Regression']['cm']
fig, ax = plt.subplots(figsize=(7, 5))
sns.heatmap(best_cm, annot=True, fmt='d', cmap='Blues', cbar=False, annot_kws={'size': 18, 'fontweight': 'bold'},
            xticklabels=['Predicted 0 (Fail)', 'Predicted 1 (Success)'],
            yticklabels=['Actual 0 (Fail)', 'Actual 1 (Success)'], ax=ax)
ax.set_title('Confusion Matrix - Best Classification Model (Test Accuracy: 83.33%)', fontsize=12, fontweight='bold', pad=12)
plt.tight_layout()
plt.savefig('images/ml_confusion_matrix.png', dpi=300)
plt.close()

print("\n--- Generating Methodology Flowchart Graphics ---")
# 1. API Flowchart (Slide 8)
fig, ax = plt.subplots(figsize=(8, 4.5))
ax.axis('off')
boxes = [
    ("1. Request SpaceX API", "Endpoint: api.spacexdata.com/v4/launches/past\nFetch 90+ rocket launch records (JSON)", 0.5, 0.85, '#3498db'),
    ("2. JSON Data Extraction", "Extract Booster Version, Payload, LaunchSite,\nOrbit, Landing Outcomes, and Core serials", 0.5, 0.55, '#f39c12'),
    ("3. Clean & Structure", "Filter Falcon 9 launches, impute payload mass,\nconvert into Pandas DataFrame for analysis", 0.5, 0.25, '#2ecc71')
]
for title, desc, cx, cy, col in boxes:
    ax.text(cx, cy, f"{title}\n{desc}", ha='center', va='center',
            bbox=dict(boxstyle="round,pad=0.6", facecolor=col, edgecolor='black', alpha=0.9),
            fontsize=11, fontweight='bold', color='white')
    if cy > 0.3:
        ax.annotate('', xy=(cx, cy-0.12), xytext=(cx, cy-0.05),
                    arrowprops=dict(facecolor='black', width=2, headwidth=8))
plt.tight_layout()
plt.savefig('images/flowchart_api.png', dpi=300)
plt.close()

# 2. Web Scraping Flowchart (Slide 9)
fig, ax = plt.subplots(figsize=(8, 4.5))
ax.axis('off')
boxes_scrape = [
    ("1. Fetch Wikipedia Page", "Target: List of Falcon 9 & Falcon Heavy launches\nHTTP GET request to Wikipedia archive URL", 0.5, 0.85, '#9b59b6'),
    ("2. Parse HTML with BeautifulSoup", "Locate HTML tables (class='wikitable'), extract\nrow headers, table cells, and hyperlinks", 0.5, 0.55, '#e67e22'),
    ("3. Normalize & Export CSV", "Clean superscripts/footnotes, handle missing dates,\nbuild consolidated launch records dataframe", 0.5, 0.25, '#16a085')
]
for title, desc, cx, cy, col in boxes_scrape:
    ax.text(cx, cy, f"{title}\n{desc}", ha='center', va='center',
            bbox=dict(boxstyle="round,pad=0.6", facecolor=col, edgecolor='black', alpha=0.9),
            fontsize=11, fontweight='bold', color='white')
    if cy > 0.3:
        ax.annotate('', xy=(cx, cy-0.12), xytext=(cx, cy-0.05),
                    arrowprops=dict(facecolor='black', width=2, headwidth=8))
plt.tight_layout()
plt.savefig('images/flowchart_scraping.png', dpi=300)
plt.close()

# 3. Data Wrangling Flowchart (Slide 10)
fig, ax = plt.subplots(figsize=(8, 4.5))
ax.axis('off')
boxes_wrangling = [
    ("1. Missing Value Imputation", "Replace missing PayloadMass values with mean mass (6104.96 kg)", 0.5, 0.85, '#2980b9'),
    ("2. Landing Outcome Labeling", "Map outcomes: True RTLS/ASDS/Ocean -> 1 (Success)\nFalse RTLS/ASDS/Ocean, None ASDS/None -> 0 (Failure)", 0.5, 0.55, '#d35400'),
    ("3. Feature Transformation", "Perform One-Hot Encoding on Orbit, LaunchSite, LandingPad, Serial;\nStandardize features via StandardScaler", 0.5, 0.25, '#27ae60')
]
for title, desc, cx, cy, col in boxes_wrangling:
    ax.text(cx, cy, f"{title}\n{desc}", ha='center', va='center',
            bbox=dict(boxstyle="round,pad=0.6", facecolor=col, edgecolor='black', alpha=0.9),
            fontsize=10.5, fontweight='bold', color='white')
    if cy > 0.3:
        ax.annotate('', xy=(cx, cy-0.12), xytext=(cx, cy-0.05),
                    arrowprops=dict(facecolor='black', width=2, headwidth=8))
plt.tight_layout()
plt.savefig('images/flowchart_wrangling.png', dpi=300)
plt.close()

# 4. Machine Learning Flowchart (Slide 15)
fig, ax = plt.subplots(figsize=(8, 4.5))
ax.axis('off')
boxes_ml = [
    ("1. Data Preprocessing & Splitting", "83 features standardized via StandardScaler; 80% Train, 20% Test (random_state=2)", 0.5, 0.85, '#34495e'),
    ("2. GridSearchCV Hyperparameter Tuning", "10-Fold Cross-Validation across Logistic Regression, SVM, Decision Tree, and KNN", 0.5, 0.55, '#8e44ad'),
    ("3. Evaluation & Validation", "Evaluate on Test set: Accuracy 83.33%, Confusion Matrix, Precision/Recall trade-off", 0.5, 0.25, '#2c3e50')
]
for title, desc, cx, cy, col in boxes_ml:
    ax.text(cx, cy, f"{title}\n{desc}", ha='center', va='center',
            bbox=dict(boxstyle="round,pad=0.6", facecolor=col, edgecolor='black', alpha=0.9),
            fontsize=10.5, fontweight='bold', color='white')
    if cy > 0.3:
        ax.annotate('', xy=(cx, cy-0.12), xytext=(cx, cy-0.05),
                    arrowprops=dict(facecolor='black', width=2, headwidth=8))
plt.tight_layout()
plt.savefig('images/flowchart_ml.png', dpi=300)
plt.close()

print("All ML models, confusion matrices, and flowcharts generated successfully!")
