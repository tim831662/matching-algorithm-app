import pandas as pd
import random
import os

# Configuration
INPUT_FILENAME = 'students_survey.csv'
OUTPUT_SCHEDULE = 'Final_Presentation_Schedule.xlsx'
OUTPUT_PRESS_POOL = 'Final_Press_Pool.xlsx'
STUDENTS_PER_GROUP = 2
GROUPS_PER_DATE = 2
MAX_STUDENTS_PER_DATE = STUDENTS_PER_GROUP * GROUPS_PER_DATE
REVIEWS_PER_STUDENT = 2

available_dates = [
    "9/8/2025", "9/10/2025", "9/15/2025", "9/17/2025", "9/22/2025",
    "9/24/2025", "9/29/2025", "10/1/2025", "10/6/2025", "10/8/2025",
    "10/13/2025", "10/15/2025", "10/20/2025", "10/22/2025", "10/27/2025",
    "10/29/2025", "11/3/2025", "11/5/2025", "11/10/2025", "11/12/2025",
    "11/17/2025", "11/19/2025", "12/1/2025", "12/3/2025", "12/8/2025"
]

def assign_presentations(students_df, dates):
    df = students_df.copy()
    df['Assigned Date'] = None
    schedule = {d: [] for d in dates}
    
    # Randomize for fairness
    df = df.sample(frac=1, random_state=42).reset_index(drop=True)
    assigned = set()
    
    # 1. Partner Matching
    for idx, student in df.iterrows():
        if idx in assigned: continue
            
        partner_name = student.get('Partner Name', 'N/A')
        if pd.isna(partner_name) or str(partner_name).strip().upper() == 'N/A': continue
        
        # Find partner index
        partner_matches = df[df['Student Name'].str.strip() == str(partner_name).strip()].index
        if len(partner_matches) == 0 or partner_matches[0] in assigned: continue
        partner_idx = partner_matches[0]
        
        # Verify mutual request
        partner_choice = df.loc[partner_idx, 'Partner Name']
        if pd.isna(partner_choice) or str(partner_choice).strip() != student['Student Name'].strip(): continue
        
        # Get all choices from both students
        s_choices = [student[f'Choice {i}'] for i in range(1, 4)]
        p_choices = [df.loc[partner_idx, f'Choice {i}'] for i in range(1, 4)]
        
        # Combine choices into one preference list
        all_choices = []
        for d in s_choices + p_choices:
            if d not in all_choices: all_choices.append(d)
            
        target_date = None
        
        # Try to fit them in one of their preferred dates
        for date in all_choices:
            if len(schedule[date]) <= MAX_STUDENTS_PER_DATE - 2:
                target_date = date
                break
        
        # Fallback
        if not target_date:
            target_date = min(schedule, key=lambda d: len(schedule[d]))
            
        # Assign both students to the target date
        schedule[target_date].extend([idx, partner_idx])
        df.at[idx, 'Assigned Date'] = target_date
        df.at[partner_idx, 'Assigned Date'] = target_date
        assigned.add(idx)
        assigned.add(partner_idx)
    
    # 2. Individual Matching
    for choice_num in range(1, 4):
        for idx, student in df.iterrows():
            if idx in assigned: continue
            
            preferred_date = student[f'Choice {choice_num}']
            if len(schedule[preferred_date]) < MAX_STUDENTS_PER_DATE:
                schedule[preferred_date].append(idx)
                df.at[idx, 'Assigned Date'] = preferred_date
                assigned.add(idx)
    
    # 3. Overflow
    for idx in df.index:
        if idx not in assigned:
            best_date = min(schedule, key=lambda d: len(schedule[d]))
            schedule[best_date].append(idx)
            df.at[idx, 'Assigned Date'] = best_date
    
    return df

def assign_press_pool(students_df, dates):
    reviews = []
    for _, student in students_df.iterrows():
        # Ensure reviewer doesn't review on their presentation day
        valid_dates = [d for d in dates if d != student['Assigned Date']]
        review_dates = random.sample(valid_dates, REVIEWS_PER_STUDENT)
        
        reviews.append({
            'Presentation Date': student['Assigned Date'],
            'Student Name': student['Student Name'],
            'Review Date 1': review_dates[0],
            'Review Date 2': review_dates[1]
        })
    return pd.DataFrame(reviews)

def save_excel(df, filename):
    try:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            for column in writer.sheets['Sheet1'].columns:
                max_len = max(len(str(cell.value)) for cell in column)
                writer.sheets['Sheet1'].column_dimensions[column[0].column_letter].width = max_len + 2
        print(f"Created: {filename}")
    except Exception as e:
        print(f"Error saving {filename}: {e}")

def main():
    if not os.path.exists(INPUT_FILENAME):
        print(f"Error: {INPUT_FILENAME} not found")
        return
    
    try:
        df = pd.read_csv(INPUT_FILENAME)
        df.columns = df.columns.str.strip()
    except Exception as e:
        print(f"Error reading CSV: {e}")
        return
    
    try:
        final_df = assign_presentations(df, available_dates)
        press_df = assign_press_pool(final_df, available_dates)
    except KeyError as e:
        print(f"Error: Missing column {e}. Ensure CSV has 'Student Name', 'Partner Name', and 'Choice 1-3'")
        return

    # Sort Chronologically
    final_df['_sort'] = pd.to_datetime(final_df['Assigned Date'], format='%m/%d/%Y', errors='coerce')
    final_df = final_df.sort_values('_sort').drop(columns=['_sort'])
    
    press_df['_sort'] = pd.to_datetime(press_df['Presentation Date'], format='%m/%d/%Y', errors='coerce')
    press_df = press_df.sort_values('_sort').drop(columns=['_sort'])
    
    # Save output
    output_cols = ['Student Name', 'Partner Name', 'Choice 1', 'Choice 2', 'Choice 3', 'Assigned Date']
    if 'Partner Name' not in final_df.columns: 
        output_cols.remove('Partner Name')

    save_excel(final_df[output_cols], OUTPUT_SCHEDULE)
    save_excel(press_df, OUTPUT_PRESS_POOL)

if __name__ == "__main__":
    main()