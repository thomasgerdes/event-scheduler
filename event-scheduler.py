import pandas as pd
from collections import defaultdict

REGISTRATIONS_FILE = 'registrations.xlsx'
PREFERENCES_FILE = 'date_survey.xlsx'

def load_registrations(file_path):
    try:
        df = pd.read_excel(file_path)
        return [(str(row['Email']).strip(), str(row['Name']).strip()) for _, row in df.iterrows() if pd.notna(row['Email']) and pd.notna(row['Name'])]
    except Exception as e:
        print(f"Error loading registrations: {e}")
        return []

def load_preferences(file_path):
    try:
        df = pd.read_excel(file_path)
        preferences = {}
        names = {}
        for _, row in df.iterrows():
            email = str(row['Email']).strip()
            name = str(row['Name']).strip()
            if pd.notna(email) and pd.notna(name):
                prefs = [f'Date {i}' for i in range(1, 4) if pd.notna(row[f'Date {i}']) and row[f'Date {i}'] == 'Yes']
                preferences[email] = prefs
                names[email] = name
        return preferences, names
    except Exception as e:
        print(f"Error loading preferences: {e}")
        return {}, {}

def distribute_participants(registrations, preferences, survey_names, max_per_slot=25):
    slots = defaultdict(list)
    waitlists = defaultdict(list)
    
    registered_emails = {email: idx for idx, (email, _) in enumerate(registrations)}
    assigned_participants = set()  # Track assigned participants
    
    # First round: distribute participants and create initial waitlists
    for email, name in registrations:
        if email in preferences and email not in assigned_participants:
            prefs = preferences[email]
            assigned = False
            for pref in prefs:
                if len(slots[pref]) < max_per_slot:
                    slots[pref].append((email, name))
                    assigned_participants.add(email)
                    assigned = True
                    break
            if not assigned:
                for pref in prefs:
                    waitlists[pref].append((email, name))
        elif email not in assigned_participants:
            # If registered but didn't participate in the survey, add only to Date 1 waitlist
            waitlists['Date 1'].append((email, name + ' *'))
    
    # Second round: handle exceptions for registered participants on waitlists
    for date, wait_list in list(waitlists.items()):
        for email, name in wait_list[:]:  # Use a copy to allow modification during iteration
            if email in registered_emails and email in preferences and email not in assigned_participants:
                prefs = preferences[email]
                least_occupied_pref = min(prefs, key=lambda x: len(slots[x]))
                slots[least_occupied_pref].append((email, name))
                assigned_participants.add(email)
                wait_list.remove((email, name))
                # Remove from other waitlists if assigned
                for other_date in waitlists:
                    if other_date != date:
                        waitlists[other_date] = [(e, n) for e, n in waitlists[other_date] if e != email]
    
    # Handle participants who only took part in the survey but didn't register
    survey_only = [(email, survey_names[email]) for email in preferences if email not in registered_emails]
    for email, name in survey_only:
        if email not in assigned_participants:
            prefs = preferences[email]
            assigned = False
            for pref in prefs:
                if len(slots[pref]) < max_per_slot:
                    slots[pref].append((email, name))
                    assigned_participants.add(email)
                    assigned = True
                    break
            if not assigned:
                for pref in prefs:
                    waitlists[pref].append((email, name))
    
    # Ensure all unassigned participants are on all their preferred waitlists
    all_participants = set(email for email, _ in registrations) | set(preferences.keys())
    unassigned_participants = all_participants - assigned_participants
    
    for email in unassigned_participants:
        if email in preferences:
            name = survey_names.get(email, next((name for e, name in registrations if e == email), None))
            if name:
                for pref in preferences[email]:
                    if (email, name) not in waitlists[pref] and (email, name + ' *') not in waitlists[pref]:
                        waitlists[pref].append((email, name))
    
    # Sort slots and waitlists based on registration order
    for date in slots:
        slots[date].sort(key=lambda x: registered_emails.get(x[0], float('inf')))
    
    for date in waitlists:
        # Sort participants without '*' first, then participants with '*', both based on registration order
        waitlists[date].sort(key=lambda x: (
            '*' in x[1],  # First sort criterion: '*' in name (False comes before True)
            registered_emails.get(x[0], float('inf'))  # Second sort criterion: registration order
        ))
    
    return slots, waitlists

def identify_contacts_to_remind(registrations, preferences):
    return [email for email, _ in registrations if email not in preferences]

def save_results(slots, waitlists, contacts_to_remind, output_file):
    try:
        # Create main data table
        main_data = {
            'Date 1': [name for _, name in slots['Date 1']],
            'Date 2': [name for _, name in slots['Date 2']],
            'Date 3': [name for _, name in slots['Date 3']],
            'Waitlist Date 1': [name for _, name in waitlists['Date 1']],
            'Waitlist Date 2': [name for _, name in waitlists['Date 2']],
            'Waitlist Date 3': [name for _, name in waitlists['Date 3']]
        }
        
        max_len = max(len(v) for v in main_data.values())
        for k in main_data:
            main_data[k] = main_data[k] + [''] * (max_len - len(main_data[k]))
        
        main_df = pd.DataFrame(main_data)
        
        # Create email data table
        email_data = {
            'Date 1 Emails': [email for email, _ in slots['Date 1']],
            'Date 2 Emails': [email for email, _ in slots['Date 2']],
            'Date 3 Emails': [email for email, _ in slots['Date 3']],
            'Waitlist Date 1 Emails': [email for email, _ in waitlists['Date 1']],
            'Waitlist Date 2 Emails': [email for email, _ in waitlists['Date 2']],
            'Waitlist Date 3 Emails': [email for email, _ in waitlists['Date 3']]
        }
        
        # Add Contact Waitlists column
        all_waitlist_emails = set()
        for waitlist in waitlists.values():
            all_waitlist_emails.update(email for email, _ in waitlist)
        email_data['Contact Waitlists'] = sorted(list(all_waitlist_emails))
        
        max_len = max(len(v) for v in email_data.values())
        for k in email_data:
            email_data[k] = email_data[k] + [''] * (max_len - len(email_data[k]))
        
        email_df = pd.DataFrame(email_data)
        
        # Save both tables to the same Excel file, but on different sheets
        with pd.ExcelWriter(output_file) as writer:
            main_df.to_excel(writer, sheet_name='Assignments', index=False)
            email_df.to_excel(writer, sheet_name='Email Addresses', index=False)
        
        print(f"Results saved to {output_file}")
    except Exception as e:
        print(f"Error saving results: {e}")

# Main function
try:
    registrations = load_registrations(REGISTRATIONS_FILE)
    preferences, survey_names = load_preferences(PREFERENCES_FILE)

    print(f"Number of registrations: {len(registrations)}")
    print(f"Number of people with date preferences: {len(preferences)}")

    # Distribution
    slots, waitlists = distribute_participants(registrations, preferences, survey_names)

    # Identify contacts to remind
    contacts_to_remind = identify_contacts_to_remind(registrations, preferences)

    # Display results
    for slot, participants in slots.items():
        print(f"{slot}: {len(participants)} participants")

    for date, waitlist in waitlists.items():
        print(f"Waitlist for {date}: {len(waitlist)} participants")

    print(f"People to contact: {len(contacts_to_remind)}")

    # Save results
    output_file = 'assignment.xlsx'
    save_results(slots, waitlists, contacts_to_remind, output_file)

    # Display detailed results
    with pd.ExcelFile(output_file) as xls:
        df_assignments = pd.read_excel(xls, 'Assignments')
        df_emails = pd.read_excel(xls, 'Email Addresses')
    
    print("\nFirst 10 rows of the assignments:")
    print(df_assignments.head(10))
    print("\nFirst 10 rows of email addresses:")
    print(df_emails.head(10))
    print("\nDistribution by date:")
    print(df_assignments.count())

    print("Code executed. The results have been saved in 'assignment.xlsx'.")
except Exception as e:
    print(f"An error occurred during execution: {e}")
