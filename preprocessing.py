import string

def clean_text(text):
    """
    Removes punctuation and converts text to lowercase.
    """
    if not isinstance(text, str):
        return ""
    
    # Remove punctuation using standard python string library
    text = text.translate(str.maketrans('', '', string.punctuation))
    # Convert to lowercase
    return text.lower()