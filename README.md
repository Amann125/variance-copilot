# Variance Copilot

Automates monthly budget vs actual commentary. Python calculates every variance and flags the material ones. An AI model drafts one line of commentary per flagged item. A verification step then checks every dollar and percent in that commentary against the source data. Anything that doesn't match gets marked for review.

**Design rule:** code does the math, the model writes the words, and a checker verifies them.

## Run it
```
pip install -r requirements.txt
python solution/variance_copilot.py --month 2026-04
```
Without an `ANTHROPIC_API_KEY` set, it uses template commentary so it still runs end to end.

## How it works
1. Load and join `data/budget.csv`, `data/actuals.csv`, and `data/notes.csv`.
2. Flag lines that are at least 5% and at least $5,000 off budget, and label each Favorable or Unfavorable.
3. Draft commentary with an AI model, constrained to the provided numbers and notes and returned as JSON.
4. Verify that every $ and % in the commentary exists in the data.
5. Write an Excel report with an expense summary tab and a flagged variances tab.

Sample data is fictional.
