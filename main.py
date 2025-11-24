from model_trainer import train_model
from predictor import predict_message

def main():
    while True:
        print("\n=== SPAMGUARD AI: SMS/EMAIL FILTER ===")
        print("1. Train AI Model (Run this first)")
        print("2. Check a Message (Spam vs Ham)")
        print("3. Exit")
        
        choice = input("Select option: ")
        
        if choice == '1':
            train_model()
        
        elif choice == '2':
            msg = input("\nEnter message to analyze: ")
            if msg.strip():
                result, confidence = predict_message(msg)
                if result == "Error":
                    print("Error: Model not found. Please train the model (Option 1) first.")
                else:
                    print(f"\n>>> Prediction: {result.upper()}")
                    print(f">>> Confidence: {confidence:.2f}%")
            else:
                print("Error: Message cannot be empty.")
                
        elif choice == '3':
            print("Exiting SpamGuard.")
            break
        else:
            print("Invalid option.")

if __name__ == "__main__":
    main()