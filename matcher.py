import pandas as pd
import random
import os
from collections import deque

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

def gale_shapley_capacity(students, prefs, dates, capacity):
    free = deque(students)
    next_idx = {s: 0 for s in students}
    accepted = {d: [] for d in dates}

    rank = {}
    n = len(dates)
    for s in students:
        rank[s] = {}
        for i, d in enumerate(prefs[s]):
            rank[s][d] = i
        for d in dates:
            if d not in rank[s]:
                rank[s][d] = n + 1

    while free:
        s = free.popleft()
        if next_idx[s] < len(prefs[s]):
            d = prefs[s][next_idx[s]]
        else:
            d = min(dates, key=lambda x: len(accepted[x]))
        next_idx[s] += 1

        if len(accepted[d]) < capacity:
            accepted[d].append(s)
        else:
            worst = max(accepted[d], key=lambda x: rank[x][d])
            if rank[s][d] < rank[worst][d]:
                accepted[d].remove(worst)
                accepted[d].append(s)
                free.append(worst)
            else:
                free.append(s)

    out = {}
    for d in dates:
        for s in accepted[d]:
            out[s] = d
    return out

def assign_presentations(students_df, dates):
    df = students_df.copy()
    df['Assigned Date'] = None

    visited = set()
    groups = []

    for idx, row in df.iterrows():
        if idx in visited:
            continue
        partner = str(row.get("Partner Name", "")).strip()
        name = row['Student Name'].strip()
        if partner and partner.upper() != "N/A":
            match = df[df['Student Name'].str.strip() == partner]
            if len(match) > 0:
                p_idx = match.index[0]
                p_partner = str(df.loc[p_idx, 'Partner Name']).strip()
                if p_partner == name:
                    groups.append((idx, p_idx))
                    visited.add(idx)
                    visited.add(p_idx)
                    continue
        groups.append((idx,))
        visited.add(idx)

    students = []
    prefs = {}

    for g in groups:
        gid = tuple(g)
        students.append(gid)
        P = []
        for member in g:
            for c in ["Choice 1", "Choice 2", "Choice 3"]:
                d = df.loc[member, c]
                if d not in P:
                    P.append(d)
        for d in dates:
            if d not in P:
                P.append(d)
        prefs[gid] = P

    assignment = gale_shapley_capacity(students, prefs, dates, GROUPS_PER_DATE)

    for g in groups:
        d = assignment[g]
        for idx in g:
            df.at[idx, 'Assigned Date'] = d

    return df

def assign_press_pool(students_df, dates):
    review_count = {d: 0 for d in dates}
    reviews = []

    shuffled_df = students_df.sample(frac=1, random_state=42).reset_index(drop=True)

    for _, student in shuffled_df.iterrows():
        assigned = student['Assigned Date']

        valid = [d for d in dates if d != assigned]

        student_reviews = []
        for _ in range(REVIEWS_PER_STUDENT):
            # Pick the date with the lowest load that isn't already chosen
            candidates = [d for d in valid if d not in student_reviews]
            best = min(candidates, key=lambda x: review_count[x])
            student_reviews.append(best)
            review_count[best] += 1

        reviews.append({
            'Presentation Date': assigned,
            'Student Name': student['Student Name'],
            'Review Date 1': student_reviews[0],
            'Review Date 2': student_reviews[1]
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

    final_df['_sort'] = pd.to_datetime(final_df['Assigned Date'], format='%m/%d/%Y', errors='coerce')
    final_df = final_df.sort_values('_sort').drop(columns=['_sort'])

    press_df['_sort'] = pd.to_datetime(press_df['Presentation Date'], format='%m/%d/%Y', errors='coerce')
    press_df = press_df.sort_values('_sort').drop(columns=['_sort'])

    output_cols = ['Student Name', 'Partner Name', 'Choice 1', 'Choice 2', 'Choice 3', 'Assigned Date']
    if 'Partner Name' not in final_df.columns:
        output_cols.remove('Partner Name')

    save_excel(final_df[output_cols], OUTPUT_SCHEDULE)
    save_excel(press_df, OUTPUT_PRESS_POOL)

if __name__ == "__main__":
    main()