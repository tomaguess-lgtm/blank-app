# Barebells @ Walmart — COG Dashboard

This repo hosts a Streamlit dashboard that reads two Excel files:
- `Barebells Supplier Summary 4.xlsx`
- `Barebells Weekly Item Tracker 4.xlsx`

## Run locally
```
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud
- Push this repo to GitHub (prefer **private**).
- On https://share.streamlit.io, New app:
  - Repository: `<your-username>/<your-repo>`
  - Branch: `main`
  - Main file path: `app.py`
- Deploy.
