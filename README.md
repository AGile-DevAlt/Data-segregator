# Event Participant Segregator

A small local web app that takes an Excel workbook with mixed participant/event data and creates a new workbook with one sheet per event, plus three day-group sheets.

## Event mapping

- **DAY 1 & 2:** Yuva Sansad, The mentalist, Drishti
- **DAY 1 ONLY:** Kalakriti, Natyalaya, JAM
- **DAY 2 ONLY:** Pitch Perfect, Reel making, Nukkad Natak, Behaviour Scientist

## Run it

1. Install Python 3.10 or newer.
2. Unzip the downloaded ZIP.
3. Open Terminal/Command Prompt in the extracted folder.
4. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

5. Start the app:

   ```bash
   streamlit run app.py
   ```

6. Your browser opens. Upload the Excel file, select the correct source worksheet, verify the column mappings, and click **Segregate participants**.
7. Download `segregated_event_participants.xlsx`.

## How it works

- It looks for event names in all cells in each row, unless you choose one or more dedicated event columns.
- A participant row mentioning multiple events is copied into each matching event sheet.
- It creates one sheet for each of the 10 events, and additional sheets for `Day 1 and 2`, `Day 1 Only`, and `Day 2 Only`.
- It tries to detect participant detail columns automatically, and lets you correct the mappings before processing.
- Rows with no detected event or no full name are skipped.
- Exact duplicate participant rows within the same event are removed.

## Important note about source layout

The app works best when each participant is represented by a row and the event name(s) for that participant appear somewhere in the same row. If your source uses a different layout (for example, participant details and event names are in separate blocks), reorganize it so each participant's event information is on the same row before uploading.

The app processes files locally on your computer when run this way; it does not send the workbook to a hosted service.
