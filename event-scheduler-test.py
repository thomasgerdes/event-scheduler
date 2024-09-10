import pandas as pd
import openpyxl
import sys
from io import StringIO

def load_data(file_path):
    """Load data from the assignment Excel file."""
    wb = openpyxl.load_workbook(file_path)
    
    def sheet_to_df(sheet):
        data = sheet.values
        cols = next(data)
        return pd.DataFrame(data, columns=cols)
    
    assignments_df = sheet_to_df(wb['Assignments'])
    emails_df = sheet_to_df(wb['Email Addresses'])
    return assignments_df, emails_df

def load_original_data(registrations_file, preferences_file):
    """Load data from original registration and preferences files."""
    registrations_df = pd.read_excel(registrations_file)
    preferences_df = pd.read_excel(preferences_file)
    return registrations_df, preferences_df

def check_all_participants_assigned(assignments_df, emails_df, registrations_df, preferences_df):
    """Check if all participants from registrations and preferences are in the assignment list."""
    all_participants = set(registrations_df['Email']) | set(preferences_df['Email'])
    assigned_participants = set()
    for column in emails_df.columns:
        assigned_participants.update(emails_df[column].dropna())
    
    missing_participants = all_participants - assigned_participants
    
    print("a) Checking: Are all participants in the assignment list?")
    if not missing_participants:
        print("   Passed: All participants are in the assignment list.")
    else:
        print(f"   Failed: {len(missing_participants)} participants are missing from the assignment list:")
        for participant in missing_participants:
            print(f"      - {participant}")

def check_all_registered_assigned_slot(emails_df, registrations_df):
    """Check if all registered participants are assigned to a slot."""
    registered_emails = set(registrations_df['Email'])
    assigned_slots = set(emails_df['Date 1 Emails']) | set(emails_df['Date 2 Emails']) | set(emails_df['Date 3 Emails'])
    unassigned_registered = registered_emails - assigned_slots
    
    print("\nb) Checking: Are all registered participants assigned to a slot?")
    if not unassigned_registered:
        print("   Passed: All registered participants are assigned to a slot.")
    else:
        print(f"   Failed: {len(unassigned_registered)} registered participants are not assigned to any slot:")
        for email in unassigned_registered:
            print(f"      - {email}")

def check_assignments_match_preferences(emails_df, preferences_df):
    """Check if assignments match the preferences stated in the survey."""
    mismatches = []
    for _, row in preferences_df.iterrows():
        email = row['Email']
        prefs = [f'Date {i}' for i in range(1, 4) if row[f'Date {i}'] == 'Yes']
        assigned_dates = []
        for date in ['Date 1', 'Date 2', 'Date 3']:
            if email in set(emails_df[f'{date} Emails']):
                assigned_dates.append(date)
            if email in set(emails_df[f'Waitlist {date} Emails']):
                assigned_dates.append(f'Waitlist {date}')
        
        if assigned_dates and not any(date.replace('Waitlist ', '') in prefs for date in assigned_dates):
            mismatches.append((email, prefs, assigned_dates))
    
    print("\nc) Checking: Do assignments match stated preferences?")
    if not mismatches:
        print("   Passed: All assignments match the stated preferences.")
    else:
        print(f"   Failed: {len(mismatches)} assignments do not match preferences:")
        for email, prefs, assigned in mismatches:
            print(f"      - {email}: Preferences {prefs}, assigned to {assigned}")

def check_double_bookings(assignments_df):
    """Check if any participant is booked for more than one slot or multiple times for the same slot in the Assignments sheet."""
    double_bookings = []
    multiple_same_slot = []
    
    for date in ['Date 1', 'Date 2', 'Date 3']:
        names = assignments_df[date].dropna().tolist()
        name_counts = pd.Series(names).value_counts()
        multiple_same_slot.extend([(name, date, count) for name, count in name_counts.items() if count > 1])
    
    all_names = assignments_df['Date 1'].tolist() + assignments_df['Date 2'].tolist() + assignments_df['Date 3'].tolist()
    all_names = [name for name in all_names if pd.notna(name)]  # Remove NaN values
    
    for name in set(all_names):
        dates = []
        if name in assignments_df['Date 1'].values:
            dates.append('Date 1')
        if name in assignments_df['Date 2'].values:
            dates.append('Date 2')
        if name in assignments_df['Date 3'].values:
            dates.append('Date 3')
        
        if len(dates) > 1:
            double_bookings.append((name, dates))
    
    print("\nd) Checking: Are there any double bookings or multiple bookings for the same slot in the Assignments sheet?")
    if not double_bookings and not multiple_same_slot:
        print("   Passed: No participant is booked for more than one slot or multiple times for the same slot.")
    else:
        if double_bookings:
            print(f"   Failed: {len(double_bookings)} participants are booked for multiple slots:")
            for name, dates in double_bookings:
                print(f"      - {name}: booked for {', '.join(dates)}")
        if multiple_same_slot:
            print(f"   Failed: {len(multiple_same_slot)} participants are booked multiple times for the same slot:")
            for name, date, count in multiple_same_slot:
                print(f"      - {name}: booked {count} times for {date}")

def run_tests(assignment_file, registrations_file, preferences_file):
    print("Starting checks on the Event Scheduler results...\n")
    
    assignments_df, emails_df = load_data(assignment_file)
    registrations_df, preferences_df = load_original_data(registrations_file, preferences_file)
    
    check_all_participants_assigned(assignments_df, emails_df, registrations_df, preferences_df)
    check_all_registered_assigned_slot(emails_df, registrations_df)
    check_assignments_match_preferences(emails_df, preferences_df)
    check_double_bookings(assignments_df)
    
    print("\nChecks completed.")

# Run the tests and save output to a file
assignment_file = 'assignment.xlsx'
registrations_file = 'registrations.xlsx'
preferences_file = 'date_survey.xlsx'

# Redirect stdout to a StringIO object
old_stdout = sys.stdout
result = StringIO()
sys.stdout = result

# Run the tests
run_tests(assignment_file, registrations_file, preferences_file)

# Restore stdout and get the output as a string
sys.stdout = old_stdout
output = result.getvalue()

# Write the output to a file
with open('test_results.txt', 'w') as f:
    f.write(output)

print("Test results have been saved to 'test_results.txt'")
