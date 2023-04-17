import pandas as pd
from collections import Counter
from tqdm import tqdm

# Variable Declarations
excel = "preprocessed_documents.xlsx"
column = "Preprocessed_Documents"
output = "tfrf_result.xlsx"

# Load in the Excel file
df = pd.read_excel(excel)

# Convert the document text to a list
docs = []
for doc in df[column]:
    if isinstance(doc, str):
        docs.append(doc.lower())
    else:
        docs.append(str(doc))

# Calculate the term frequencies for each document
doc_term_freq = []
for doc in docs:
    # Split the document into individual terms
    terms = doc.split()
    # Count the occurrences of each term
    term_counts = Counter(terms)
    # Calculate the term frequency for each term
    term_freq = {
        term: count / len(terms)
        for term, count in term_counts.items()
    }
    # Add the term frequencies for the document to the list
    doc_term_freq.append(term_freq)

# Calculate the document frequency for each term
doc_freq = Counter()
for doc in doc_term_freq:
    doc_freq.update(doc.keys())

# Calculate the relative frequency for each term
total_terms = sum(doc_freq.values())
term_rf = {term: count / total_terms for term, count in doc_freq.items()}

# Calculate the TF-RF scores for each document
doc_tf_rf = []
for i, term_freq in tqdm(enumerate(doc_term_freq), total=len(doc_term_freq)):
    tf_rf = {term: freq * term_rf[term] for term, freq in term_freq.items()}
    doc_tf_rf.append(tf_rf)

    # Add a progress percentage to the console output
    progress = int(((i + 1) / len(doc_term_freq)) * 100)
    print(f'Progress: {progress}%')

# Convert the TF-RF scores to a DataFrame
tf_rf_df = pd.DataFrame(doc_tf_rf)

# Add the document number column as the first column
doc_nums = [f'Doc {i+1}' for i in range(len(docs))]
tf_rf_df.insert(0, 'Doc Number', doc_nums)

# Fill in missing values with 0
tf_rf_df.fillna(0, inplace=True)

# Export the results to a new Excel file
tf_rf_df.to_excel(output, index=False)