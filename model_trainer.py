import pandas as pd
import joblib
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.pipeline import make_pipeline
from sklearn.model_selection import train_test_split
from preprocessing import clean_text
import os

DATA_PATH = os.path.join(os.path.dirname(__file__), '../data/spam_data.csv')
MODEL_PATH = os.path.join(os.path.dirname(__file__), '../models/spam_classifier.pkl')

def train_model():
    print("--- Training Model ---")
    try:
        # 1. Load Data
        df = pd.read_csv(DATA_PATH)
        
        # 2. Preprocess
        df['text'] = df['text'].apply(clean_text)
        
        X = df['text']
        y = df['label']
        
        # 3. Create a Pipeline (Vectorizer + Classifier)
        # CountVectorizer turns words into numbers. MultinomialNB is the AI.
        model = make_pipeline(CountVectorizer(), MultinomialNB())
        
        # 4. Train
        model.fit(X, y)
        print("Model trained successfully.")
        
        # 5. Save Model
        os.makedirs(os.path.dirname(MODEL_PATH), exist_ok=True)
        joblib.dump(model, MODEL_PATH)
        print(f"Model saved to {MODEL_PATH}")
        
    except FileNotFoundError:
        print("Error: data/spam_data.csv not found. Please run create_dummy_data.py first.")
    except Exception as e:
        print(f"An error occurred: {e}")