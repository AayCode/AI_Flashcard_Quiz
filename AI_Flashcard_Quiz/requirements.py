import nltk

resources = [
    "punkt",
    "punkt_tab",
    "stopwords",
    "wordnet",
    "omw-1.4",
    "averaged_perceptron_tagger",
    "maxent_ne_chunker",
    "words"
]

for resource in resources:
    print(f"Downloading: {resource} ...")
    nltk.download(resource)

print("✅ All downloads complete!")