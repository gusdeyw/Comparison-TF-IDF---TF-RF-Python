import pandas as pd
import nltk
import re
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize

# Download the Indonesian and English stopwords corpus from NLTK
nltk.download('stopwords')
nltk.download('punkt')
nltk.download('words')

# Read the input Excel file containing the documents
input_file = 'input_documents.xlsx'
df = pd.read_excel(input_file)

# Get the document column as a list
documents = df['Document'].tolist()

# Define the Indonesian and English stopwords lists
stop_words_indo = set(stopwords.words('indonesian'))
stop_words_en = set(stopwords.words('english'))


# Define a function to perform text preprocessing
def preprocess_text(text):
    # Convert the text to lowercase
    text = text.lower()

    # Remove non-alphanumeric characters and extra whitespaces
    text = re.sub(r'\W+', ' ', text)
    text = re.sub(r'\s+', ' ', text)

    # Tokenize the text
    words = word_tokenize(text)

    # Filter out Indonesian and English stopwords
    filtered_words = []
    for word in words:
        if word not in stop_words_indo and word not in stop_words_en:
            filtered_words.append(word)

    # Join the remaining words into a single string
    text = ' '.join(filtered_words)

    return text


# Preprocess each document in the set
preprocessed_documents = []
for doc in documents:
    preprocessed_doc = preprocess_text(doc)
    preprocessed_documents.append(preprocessed_doc)

# Create a pandas DataFrame to store the preprocessed documents
df_preprocessed = pd.DataFrame(preprocessed_documents,
                               columns=['Preprocessed_Documents'])

# Export the DataFrame to a new Excel file
output_file = 'preprocessed_documents.xlsx'
df_preprocessed.to_excel(output_file, index=False)

# Print a message indicating that the export was successful
print(f"Preprocessed documents exported to {output_file}.")