# SpaceX Falcon 9 First Stage Landing Prediction

A data science capstone project from the IBM Data Science Professional Certificate on Coursera. The goal is to predict whether the SpaceX Falcon 9 first stage will successfully land after launch — information that directly impacts the cost of a launch (~$62M for SpaceX vs. ~$165M for competitors who cannot reuse boosters).

## Project Overview

This project follows a full data science workflow, from raw data collection through to an interactive dashboard and machine learning predictions.

## Notebooks & Files

| File | Description |
|------|-------------|
| `jupyter-labs-spacex-data-collection-api.ipynb` | Collects SpaceX launch data using the SpaceX REST API |
| `jupyter-labs-webscraping.ipynb` | Scrapes additional launch records from Wikipedia using BeautifulSoup |
| `labs-jupyter-spacex-Data wrangling.ipynb` | Cleans and prepares the dataset, creates the landing outcome labels |
| `jupyter-labs-eda-sql-coursera_sqllite.ipynb` | Exploratory data analysis using SQL queries on a SQLite database |
| `edadataviz.ipynb` | Visual EDA with Matplotlib and Seaborn — launch success trends, payload vs. orbit, etc. |
| `lab_jupyter_launch_site_location.ipynb` | Geospatial analysis of launch sites using Folium interactive maps |
| `SpaceX_Machine Learning Prediction_Part_5.ipynb` | Trains and evaluates four classifiers (Logistic Regression, SVM, Decision Tree, KNN) using GridSearchCV |
| `spacex-dash-app.py` | Interactive Plotly Dash dashboard for exploring launch records |

## Machine Learning Results

All four models were evaluated on the same test set and achieved the same test accuracy:

| Model | Test Accuracy |
|-------|--------------|
| Logistic Regression | 83.3% |
| Support Vector Machine | 83.3% |
| Decision Tree | 83.3% |
| K-Nearest Neighbors | 83.3% |

The Decision Tree achieved the highest cross-validation score (87.3%) during hyperparameter tuning.

## Running the Dashboard

**Install dependencies:**
```bash
pip install pandas dash plotly
```

**Launch the app:**
```bash
python spacex-dash-app.py
```

Then open `http://127.0.0.1:8050` in your browser. The dashboard includes:
- A dropdown to filter by launch site
- A pie chart showing launch success rates
- A payload range slider
- A scatter plot correlating payload mass with launch outcome

## Tech Stack

- **Data:** SpaceX REST API, Wikipedia web scraping
- **Analysis:** Python, Pandas, NumPy, SQL (SQLite)
- **Visualization:** Matplotlib, Seaborn, Folium, Plotly
- **Machine Learning:** scikit-learn (GridSearchCV, Logistic Regression, SVM, Decision Tree, KNN)
- **Dashboard:** Plotly Dash
