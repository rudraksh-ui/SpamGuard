import pandas as pd
import os

def create_data():
    # 1. Define the correct path based on WHERE this script file is located
    # This ensures it always finds the 'data' folder inside SpamGuard_Project
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(current_dir)
    data_path = os.path.join(project_root, 'data', 'spam_data.csv')
    data_dir = os.path.dirname(data_path)

    # 2. Create the data
    data = {
        'text': [
            'Win a free iPhone now! Click here.',
            'Hey, are we still meeting for lunch?',
            'Urgent! Your bank account is locked. Verify now.',
            'Mom called, she wants you to buy milk.',
            'Congratulations, you won the lottery!',
            'Please review the attached project report.',
            'Free entry in 2 a wkly comp to win FA Cup final tkts',
            'I will be late for the class today.',
            'XXXMobileMovieClub: To use your credit, click the WAP link',
            'Can you send me the notes for Python?'
        ],
        'label': ['spam', 'ham', 'spam', 'ham', 'spam', 'ham', 'spam', 'ham', 'spam', 'ham']
    }
    df = pd.DataFrame(data)
    
    # 3. Create folder if it doesn't exist
    os.makedirs(data_dir, exist_ok=True)
    
    # 4. Save
    df.to_csv(data_path, index=False)
    print(f"SUCCESS: Dataset created at: {data_path}")

if __name__ == "__main__":
    create_data()