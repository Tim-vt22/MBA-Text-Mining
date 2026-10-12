# Customer Review Text Mining – Women's Clothing

Group assignment for the course *Text Mining / Text Analytics* (Hochschule Ludwigshafen, Dr. Richard Gruss).
Task: analyse the customer reviews in `clothing_reviews.xlsx` with text mining methods and present the findings to management.

## Contents

| File | What it is |
|---|---|
| `customer_review_text_mining.ipynb` | The analysis (Jupyter notebook, with outputs) |
| `Customer_Review_Text_Mining.pptx` | Management presentation (13 slides) |
| `build_presentation.py` | Builds the presentation from the notebook outputs |
| `clothing_reviews.xlsx` | Data: 23,490 reviews |
| `figures/` | Charts exported by the notebook |
| `results/key_numbers.json` | Key figures exported by the notebook (used by the presentation) |

## Methods (following lecture 1, slides 21–23 and 33–49)

1. Corpus construction: title + review text, 22,641 documents
2. From text to numbers: review length, word length, vocabulary, Flesch readability
3. NLP preprocessing: tokenization, stop-word removal, lemmatization (NLTK)
4. Word and n-gram frequencies, word clouds, sizing n-grams
5. Word co-occurrence
6. Document-term matrix, TF-IDF and log-odds: terms that separate negative from positive reviews
7. Sentiment analysis with VADER, validated against the star rating
8. Topic modeling with LDA (scikit-learn) on 1–3★ reviews
9. Topics by department and age group

## Key findings

* High satisfaction: average 4.18★, 82% recommend.
* 20% of reviews report a size mismatch; items tend to run large.
* LDA finds 8 complaint themes in critical reviews – led by unflattering cut & shape (24%) and sizing (20%).
* Terms most typical for bad reviews: *poor quality*, *not worth*, *unflattering*, *cheap*, *shapeless*.
* VADER sentiment follows the rating (ρ = 0.45) but scores 70% of 1–2★ reviews as positive ("*I wanted to love this…*").

## Reproduce

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows
pip install -r requirements.txt
jupyter nbconvert --to notebook --execute --inplace customer_review_text_mining.ipynb
python build_presentation.py
```
