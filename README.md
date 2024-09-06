# Event Scheduler

## Description
This Python script is designed to allocate participants to multiple event dates based on their preferences and registration status. It's ideal for organizing events with limited capacity per date while respecting participant preferences.

## Key Features
- Loads registrations and date preferences from Excel files
- Distributes participants across multiple dates (up to 25 per date)
- Prioritizes registered participants who completed the date survey
- Creates waitlists for each date
- Identifies participants who need to be contacted (registered but didn't complete the survey)
- Exports results to an Excel file

## Requirements
- Python 3.6+
- `pandas` library

## Usage
1. Prepare two Excel files:
   - `registrations.xlsx`: Contains registrations (columns: 'Email', 'Name')
   - `date_survey.xlsx`: Contains date preferences (columns: 'Email', 'Name', 'Date 1', 'Date 2', 'Date 3')
2. Place these files in the same directory as the script
3. Run the script:
   ```bash
   python event_scheduler.py
   ```
4. Results will be saved in `assignment.xlsx`

## Note
This version maintains a limit of 25 participants per date and creates separate waitlists for each date. Participants who only registered but didn't complete the survey are automatically added to the Date 1 waitlist.
