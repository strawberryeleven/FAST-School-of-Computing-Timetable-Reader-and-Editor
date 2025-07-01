# Timetable Reader for FAST School of Computing

## Description

**Timetable Reader** is a Python-based desktop application built for **FAST-NUCES** to parse, edit, and export semester timetable data from Excel files. The system supports **role-based access** for different stakeholders, allowing **academic officers, HODs, lab committees, and students** to interact with timetable data according to their permissions.

The project was developed as part of the *Software Requirements Engineering* course. Functional and non-functional requirements were gathered through structured meetings with the course instructor.

---

## Table of Contents

- [Installation](#installation)  
- [Usage](#usage)  
- [Features](#features)  
- [Technologies Used](#technologies-used)  
- [Contributors](#contributors)  
- [License](#license)

---

## Installation

1. **Clone the repository:**

```bash
git clone https://github.com/yourusername/timetable-reader.git
cd timetable-reader
```

2. **Install dependencies:**

```bash
pip install pandas openpyxl
```

3. **Run the parser:**

```bash
python timetable_parser.py
```

---

## Usage

1. Place the **Excel timetable file (.xlsx)** in the project directory.

2. Run the script to extract and process the data.

3. Output will be generated as:
   - `processedTimeTableCSV.csv`
   - `processedTimeTableEXCEL.xlsx`

---

## Features

- Extracts and cleans timetable data from FAST Excel sheets
- Role-based access control for various users:
  - **Academic Officers**: Edit venues, timings, and days
  - **HOD/HOS**: Manage instructors and courses
  - **Lab Committee**: Edit lab instructor assignments
  - **Students**: View-only access
- Outputs updated data in `.csv` and `.xlsx` formats
- Input data validation and merging with course metadata

---

## Technologies Used

- Python  
- pandas  
- openpyxl  
- tkinter  

---

## Contributors

- Muhammad Ibrahim Zafar
- Hafiz Ibrahim Gul Butt  
- Shaheer Hashmi  
- Zeeshan Ahmad  
- Abdul Haseeb  

---

## License

This project is for educational purposes only.
