# Student Allotment for Webcasting

Simple web app for student polling-station registration and admin Excel export.

## Run

```bash
pip install -r requirements.txt
python app.py
```

Open: http://127.0.0.1:5000

## Admin

- URL: http://127.0.0.1:5000/admin/login
- Username: `admin`
- Password: `admin@chai$$1`

## Features

- Students register without login (roll number: 10 alphanumeric)
- College and Branch (AIML) are fixed
- Contact number auto-fills from master list; Edit unlocks it
- Changed contacts are highlighted yellow in Excel export
- Cascading dropdowns: Mandal → Gram Panchayat/Ward → Polling Station No → Building
- Admin can view all registrations and export Excel
