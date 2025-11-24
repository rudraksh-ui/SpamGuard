# SpamGuard AI 🛡️

**Course:** Fundamentals in AI and ML
**Project:** Spam Filtering System using Naive Bayes
**Student Name:** Vatsal Kirti Srivastava
**Reg No:** 25BSA10076

## 📝 Overview
SpamGuard is a Python-based AI tool that detects spam messages. It utilizes Natural Language Processing (NLP) techniques to vectorize text and a Multinomial Naive Bayes classifier to predict if a message is malicious (Spam) or safe (Ham).

## ✨ Features
* **NLP Preprocessing:** auto-cleans punctuation and casing.
* **ML Classification:** Uses Scikit-Learn's Naive Bayes.
* **Model Persistence:** Saves trained models using Joblib.
* **CLI Interface:** Simple, menu-driven user interaction.

## 🛠️ Tech Stack
* Python 3.x
* Scikit-Learn (Machine Learning)
* Pandas (Data Handling)
* Joblib (Model Saving)

## 🚀 How to Run
1.  **Install Requirements:**
    `pip install pandas scikit-learn joblib`
2.  **Generate Data:**
    `python src/create_dummy_data.py`
3.  **Run System:**
    `python src/main.py`
