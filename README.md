# Work Schedule Automation

Automatically fetches work schedules from Gmail and adds them to Google Calendar.

## Quick Setup

1. **Install dependencies:**
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```

2. **Get Google API credentials:**
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Enable Gmail API and Calendar API
   - Create OAuth credentials → Download as `credentials.json`

3. **Configure:**
   Edit `main.py`:
   ```python
   name = "Your Name"  # Line 17
   employer_email = "your.employer@company.com"  # Line 218
   ```

4. **Run:**
   ```bash
   python3 main.py
   ```

## What It Does

- Searches Gmail for Excel schedules from your employer
- Finds your shifts and creates calendar events
- Handles time formats like `5:30-CL` and `6-CL`

## Troubleshooting

- **"No shifts found"**: Check your name matches the Excel file exactly
- **"ModuleNotFoundError"**: Activate virtual environment first
- **Auth issues**: Delete `token.json` and run again

## Security

Never commit `credentials.json` or `token.json` - they contain sensitive data.