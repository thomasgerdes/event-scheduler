# Event Scheduler

## Description
This Python project consists of two main scripts: an event scheduler for allocating participants to multiple event dates based on their preferences and registration status, and a test script to verify the correctness of the allocations. It's ideal for organizing events with limited capacity per date while respecting participant preferences.

## Key Features
- Loads registrations and date preferences from Excel files
- Distributes participants across multiple dates (up to 25 per date)
- Prioritizes registered participants who completed the date survey
- Creates waitlists for each date
- Identifies participants who need to be contacted (registered but didn't complete the survey)
- Exports results to an Excel file
- Includes a comprehensive test script to validate the event scheduler's output

## Requirements
- Python 3.6+
- `pandas` library
- `openpyxl` library

## Usage

### Event Scheduler
1. Prepare two Excel files:
   - `registrations.xlsx`: Contains registrations (columns: 'Email', 'Name')
   - `date_survey.xlsx`: Contains date preferences (columns: 'Email', 'Name', 'Date 1', 'Date 2', 'Date 3')
2. Place these files in the same directory as the script
3. Run the script:
   ```bash
   python event_scheduler.py
   ```
4. Results will be saved in `assignment.xlsx`

### Test Script
1. Ensure `assignment.xlsx`, `registrations.xlsx`, and `date_survey.xlsx` are in the same directory as the test script
2. Run the test script:
   ```bash
   python event_scheduler_test.py
   ```
3. Test results will be saved in `test_results.txt`

## Notes
- This version maintains a limit of 25 participants per date and creates separate waitlists for each date.
- Participants who only registered but didn't complete the survey are automatically added to the Date 1 waitlist.
- The test script checks for:
  a) All participants being in the assignment list
  b) All registered participants being assigned to a slot
  c) Assignments matching stated preferences
  d) No double bookings or multiple bookings for the same slot

## Updates (as of September 11, 2024)
- Added a comprehensive test script (`event_scheduler_test.py`) to validate the output of the event scheduler
- Test results are now saved to a text file for easy review
- Improved error handling and data validation in both scripts
- Enhanced the check for double bookings to include multiple bookings for the same slot

For any issues or suggestions, please open an issue in the GitHub repository.
