# Google Sheets Custom Script

This script enhances Google Sheets by adding a custom menu with two functions: scheduling consultations and recalculating financial margins. It automates event creation and ensures proper scheduling while preventing conflicts.

## Features

### 1. Custom Menu
When the spreadsheet opens, the script creates a menu named **"Custom Scripts"** with two options:
- **Buat Jadwal Konsul**: Manually schedules a medical consultation event.
- **Hitung Margin**: Recalculates financial formulas.

### 2. Scheduling Consultation Events
- Collects existing dates from the sheet.
- Identifies the latest valid consultation date.
- Prevents scheduling conflicts by checking for duplicate or nearby dates.
- Creates Google Calendar events with email invitations and reminders.
- Updates event status in the spreadsheet.

### 3. Margin Calculation
- Recalculates specific financial formulas when the "HITUNG ULANG" flag is found.
- Updates calculated values in multiple columns.
- Marks completed calculations as "SELESAI".
- Displays a confirmation message after updates.

## Usage
1. Open the Google Sheet.
2. Navigate to the **"Custom Scripts"** menu.
3. Select **"Buat Jadwal Konsul"** to schedule a consultation.
4. Select **"Hitung Margin"** to update financial calculations.

## Requirements
- Google Sheets with appropriate column structure.
- Google Calendar permissions for event creation.

This script simplifies scheduling and financial calculations, making it a powerful tool for managing consultations and tracking margins efficiently.
