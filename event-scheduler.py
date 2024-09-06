import pandas as pd
from collections import defaultdict

def load_registrations(file_path):
    df = pd.read_excel(file_path)
    return [(row['Email'], row['Name']) for _, row in df.iterrows()]

def load_preferences(file_path):
    df = pd.read_excel(file_path)
    preferences = {}
    names = {}
    for _, row in df.iterrows():
        email = row['Email']
        name = row['Name']
        prefs = [f'Date {i}' for i in range(1, 4) if row[f'Date {i}'] == 'Yes']
        preferences[email] = prefs
        names[email] = name
    return preferences, names

def distribute_participants(registrations, preferences, survey_names, max_per_slot=25):
    slots = defaultdict(list)
    waitlists = defaultdict(list)
    assigned_participants = set()
    
    # First, handle participants from the registration list
    for email, name in registrations:
        if email in preferences:
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
        else:
            # If registered but didn't participate in the survey, add directly to Date 1 waitlist
            waitlists['Date 1'].append((email, name + ' *'))
    
    # Then, distribute participants who only took part in the survey
    survey_only = [(email, survey_names[email]) for email in preferences if email not in dict(registrations)]
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
    
    # Sort waitlists
    registration_order = {email: idx for idx, (email, _) in enumerate(registrations)}
    for date in waitlists:
        waitlists[date] = sorted(waitlists[date], 
                                 key=lambda x: (x[0] not in preferences, 
                                                registration_order.get(x[0], float('inf'))))
    
    return slots, waitlists

def identify_contacts_to_remind(registrations, preferences):
    return [email for email, _ in registrations if email not in preferences]

def save_results(slots, waitlists, contacts_to_remind, output_file):
    data = {
        'Date 1': [name for _, name in slots['Date 1']],
        'Date 2': [name for _, name in slots['Date 2']],
        'Date 3': [name for _, name in slots['Date 3']],
        'Waitlist Date 1': [name for _, name in waitlists['Date 1']],
        'Waitlist Date 2': [name for _, name in waitlists['Date 2']],
        'Waitlist Date 3': [name for _, name in waitlists['Date 3']],
        'Contact': contacts_to_remind
    }
    
    max_len = max(len(v) for v in data.values())
    for k in data:
        data[k] = data[k] + [''] * (max_len - len(data[k]))
    
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)

# Main function
registrations = load_registrations('registrations.xlsx')
preferences, survey_names = load_preferences('date_survey.xlsx')

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
print(f"\nResults have been saved in '{output_file}'.")

# Display detailed results
df = pd.read_excel(output_file)
print("\nFirst 10 rows of the assignment:")
print(df.head(10))
print("\nDistribution by date:")
print(df.count())

print("Code executed. The results have been saved in 'assignment.xlsx'.")
