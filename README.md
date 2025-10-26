# Timetabling with OR-Tools

A lightweight Python project that uses **Google OR-Tools** to generate an optimal class schedule for students and teachers based on availability, qualifications, and course demand. It demonstrates constraint-based optimization for real-world scheduling problems using a small Excel dataset.


## Setup Instructions

### 1. Create and activate a virtual environment

#### macOS/Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
```

#### Windows

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the solver

Place your input Excel file `timetabling_input.xlsx` in the same directory as `timetabling_optimization.py`.
Then run:

```bash
python timetabling_optimization.py
```

The program will read from **timetabling_input.xlsx**, compute the optimal timetable, and print assignments for both students and teachers to the console.

---

## Project Structure

```
project_root/
├── solver.py
├── requirements.txt
└── timetabling_input.xlsx
```