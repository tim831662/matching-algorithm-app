# CS3604 Presentation Matching Algorithm

Assigns student presentation dates and press pool reviews using the Gale-Shapley stable matching algorithm.

## What It Does

1. Matches ~100 students to 25 presentation dates based on preferences
2. Assigns each student 2 review dates (not on their own presentation day)
3. Supports partner requests
4. Ensures balanced distribution (4 students per date)

## Requirements

- Python 3.7+
- pandas
- openpyxl

Install dependencies:
```bash
pip install pandas openpyxl
```

## Files

- `matcher.py` - Main matching algorithm
- `create_csv.py` - Generates sample data
- `students_survey.csv` - Input file with student preferences
- `Final_Presentation_Schedule.xlsx` - Output with presentation assignments
- `Final_Press_Pool.xlsx` - Output with review assignments

## How to Use

### 1. Create Input File

Create `students_survey.csv` with these columns:
```csv
Student Name,Choice 1,Choice 2,Choice 3,Partner Name
Alpha Student,9/8/2025,9/10/2025,9/15/2025,Beta Student
Beta Student,9/8/2025,9/22/2025,10/1/2025,Alpha Student
Gamma Student,9/8/2025,9/10/2025,10/6/2025,N/A
```

**Notes:**
- Date format: M/D/YYYY
- Partner requests must be mutual
- Use "N/A" for students without partners

### 2. Run the Program

Generate sample data (optional):
```bash
python create_csv.py
```

Run the matcher:
```bash
python matcher.py
```

### 3. Check Output

Two Excel files are created:
- `Final_Presentation_Schedule.xlsx` - Shows assigned presentation dates
- `Final_Press_Pool.xlsx` - Shows assigned review dates

## Configuration/Scalability

Edit these values in `matcher.py`:
```python
STUDENTS_PER_GROUP = 2        # Students per presentation group
GROUPS_PER_DATE = 2           # Groups per date
REVIEWS_PER_STUDENT = 2       # Reviews each student does
```

To change available dates, edit the `available_dates` list in `matcher.py`.