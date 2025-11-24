import joblib
import os
from preprocessing import clean_text

MODEL_PATH = os.path.join(os.path.dirname(__file__), '../models/spam_classifier.pkl')

def predict_message(message):
    try:
        # Load the model
        model = joblib.load(MODEL_PATH)
        
        # Clean input
        cleaned_msg = clean_text(message)
        
        # Predict
        prediction = model.predict([cleaned_msg])[0]
        # Get confidence (probability)
        probability = model.predict_proba([cleaned_msg]).max() * 100
        
        return prediction, probability
        
    except FileNotFoundError:
        return "Error", 0