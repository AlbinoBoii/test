import os
import requests
import schedule
import time
import threading
import telebot
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime
from flask import Flask, request

# Initialize the bot with the API key
API_KEY = "7080788214:AAGItH2x8AFszKuf4hw11o0opVPSgxEFhzM"
bot = telebot.TeleBot(API_KEY)

app = Flask(__name__)

# Google Sheets setup
SERVICE_ACCOUNT_FILE = r'credentials.env'
SAMPLE_SPREADSHEET_ID = "1EiHwDdD6dUYTRcU7UHNcm61E8X6of2SYVCWSYfWFBN0"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)

DEFAULT_RANGE = "A2:AH17"

# Set webhook URL - replace <your_render_url> with your actual Render URL
WEBHOOK_URL = f"https://projects-s5cf.onrender.com/{API_KEY}"

# Discord webhook for sending debug logs
DISCORD_WEBHOOK_URL = "https://discord.com/api/webhooks/1303046966514810954/Tp7DYg-Fieec2e8nIVCrgSvN_rO5rQGZubcMdSNzLXgGmcY4RHp5WqGktG7bVC_RxBW-"

# Dictionary to track individual user sessions
user_sessions = {}

# Functions remain unchanged, except where marked for session-specific updates.

def send_debug_to_discord(message):
    """Send debugging messages to Discord."""
    if DISCORD_WEBHOOK_URL:
        data = {"content": message}
        try:
            response = requests.post(DISCORD_WEBHOOK_URL, json=data)
            if response.status_code != 204:
                print(f"Failed to send to Discord: {response.status_code} - {response.text}")
        except Exception as e:
            print(f"Error sending message to Discord: {e}")

# Ping server function remains unchanged.
def ping_server():
    url = "https://projects-s5cf.onrender.com"  # Use actual server URL
    try:
        response = requests.get(url)
        print(f"Pinged server, status code: {response.status_code}")
    except requests.RequestException as e:
        print(f"Failed to ping server: {e}")

# Schedule the job to run every 14 minutes
schedule.every(14).minutes.do(ping_server)

# Background thread for running the ping scheduler
def run_scheduler():
    while True:
        schedule.run_pending()
        time.sleep(1)

@app.route('/health-check', methods=['GET'])
def health_check():
    send_debug_to_discord("Health check - bot responded!")  # Send to Discord
    return "Bot is responsive!", 200

@app.route('/setwebhook', methods=['GET'])
def set_webhook():
    success = bot.set_webhook(url=WEBHOOK_URL)
    message = "Webhook setup successful!" if success else "Webhook setup failed."
    send_debug_to_discord(message)  # Log to Discord
    return message, 200 if success else 500

@app.route('/' + API_KEY, methods=['POST'])
def get_message():
    json_str = request.get_data().decode('UTF-8')
    update = telebot.types.Update.de_json(json_str)
    bot.process_new_updates([update])
    return "!", 200

@app.route('/')
def index():
    send_debug_to_discord("Bot is running!")  # Send status to Discord
    return "Bot is running!", 200

# Helper function to list sheet names
def get_sheet_names(service, spreadsheet_id):
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        return [sheet['properties']['title'] for sheet in spreadsheet['sheets']]
    except HttpError as err:
        error_message = f"An error occurred: {err}"
        send_debug_to_discord(error_message)
        return []

# Function to fetch and clean sheet data
def fetch_sheet_data(sheet_range):
    sheet = service.spreadsheets()
    try:
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=sheet_range).execute()
        values = result.get('values', [])
        if not values:
            print("No data found in the sheet.")
            return pd.DataFrame()  # Return an empty DataFrame if no data is found

        cleaned_values = [[cell.replace('\n', ' ') if cell else cell for cell in row] for row in values]
        df = pd.DataFrame(cleaned_values)
        print("Fetched DataFrame Columns:", df.columns)
        print("Fetched DataFrame Preview:\n", df.head())
        return df
    except HttpError as e:
        print(f"Error fetching data from Google Sheets: {e}")
        return pd.DataFrame()  # Return an empty DataFrame on error

@bot.message_handler(commands=['start'])
def start(message):
    chat_id = message.chat.id
    user_sessions[chat_id] = {"sheet_name": None, "range_name": None}  # Initialize session for the user
    sheet_names = get_sheet_names(service, SAMPLE_SPREADSHEET_ID)
    if sheet_names:
        sheet_list = "\n".join(f"{i + 1}: {name}" for i, name in enumerate(sheet_names))
        bot.reply_to(message, f"Welcome! \n\nPlease select a spreadsheet:\n\n{sheet_list}")
        bot.register_next_step_handler(message, handle_sheet_selection)
    else:
        bot.reply_to(message, "No sheets found in the spreadsheet.")

def handle_sheet_selection(message):
    chat_id = message.chat.id
    user_session = user_sessions.get(chat_id, {})
    sheet_names = get_sheet_names(service, SAMPLE_SPREADSHEET_ID)

    try:
        sheet_index = int(message.text) - 1
        if 0 <= sheet_index < len(sheet_names):
            user_session["sheet_name"] = sheet_names[sheet_index]
            user_session["range_name"] = f"{user_session['sheet_name']}!{DEFAULT_RANGE}"
            bot.reply_to(message, f"Sheet '{user_session['sheet_name']}' selected. Now, please enter the day of the month (e.g., '9').")
            bot.register_next_step_handler(message, fetch_roster_for_day)
        else:
            bot.reply_to(message, "Invalid selection. Please enter a number corresponding to a sheet.")
            bot.register_next_step_handler(message, handle_sheet_selection)
    except ValueError:
        bot.reply_to(message, "Please enter a valid number corresponding to a sheet.")
        bot.register_next_step_handler(message, handle_sheet_selection)



# Keywords that indicate the cell should be treated as empty hence ignored, see ah the thing is only the top left hand coner of a merged cell show any value and the rest of the merged cell is treated as empty by the api so i had to use this hack to get around it 
keywords_to_ignore = ["M1", "M2", "M3", "M4", "M-1", "M-2", "M-3", "M-4", "C1", "C-1", "C5", "C-5",  "M3!", "OJT", "MA", "DO", "DOIL", "OIL", "SMOKE COVER", "AMPT", "OFF"]
keywords_for_duty_personnel = ["M1", "M2", "M3", "M4", "C1", "C5", "M-1", "M-2", "M-3", "M-4", "C-1", "C-5"]
keywords_for_present_personnel = ["M1", "M2", "M3", "M4", "C1", "C5", "M-1", "M-2", "M-3", "M-4", "C-1", "C-5", "M3!", "OJT" ]

# Core function to print roster based on day
def print_roster_for_day(df, day_of_month):
    column_index = 3 + (day_of_month - 1)
    if column_index >= len(df.columns):
        return "Invalid day of the month."

    personnel_names = df.iloc[4:15, 1]  # Personnel names are in a fixed column

    # Function to find duty by looking left until a non-empty cell is found
    def find_duty(row, col):
        while col >= 0:
            duty = df.iat[row, col]
            if duty:  # Check if the cell is not empty
                if any(keyword in duty for keyword in keywords_to_ignore):
                    return ""  # Treat this duty as blank
                else:
                    return duty  # Return the first valid duty found
            col -= 1
        return ""  # Return blank if no valid duty is found

    # Loop over the target rows to check each personnel’s duty
    duty_for_day = []
    for row in range(4, 15):  # Rows 4 to 14 contain the duty information
        duty = df.iat[row, column_index] if df.iat[row, column_index] else find_duty(row, column_index - 1)
        
        # Replace "x" with a blank
        if duty.lower().strip() == "x":
            duty = ""
        
        duty_for_day.append(duty)

    # Create a DataFrame with personnel names and corresponding duties
    duty_pairing = pd.DataFrame({
        'Personnel Names': personnel_names,
        'Duty for Day': duty_for_day
    }).drop_duplicates(subset='Personnel Names')

    # Identify "OWADIO" or "ORD" personnel
    ord_personnel = duty_pairing[duty_pairing['Duty for Day'].str.contains("OWADIO|ORD", na=False)]['Personnel Names'].tolist()

    # Recalculate `total_medics` by excluding personnel in `ord_personnel`
    filtered_medics = duty_pairing[~duty_pairing['Personnel Names'].isin(ord_personnel)]
    total_medics = len(filtered_medics) - 1  # Exclude CPC if necessary

    print("Number of Personnel in Duty Pairing:", total_medics)  # Debug output

    # Separate duty and non-duty personnel based on the defined keywords
    duty_personnel = duty_pairing[duty_pairing['Duty for Day'].str.contains('|'.join(keywords_for_duty_personnel), na=False)]
    non_duty_personnel = duty_pairing[~duty_pairing['Duty for Day'].str.contains('|'.join(keywords_for_duty_personnel), na=False)]

    # Return ord_personnel so it can be used in other functions
    return duty_pairing, duty_personnel, non_duty_personnel, total_medics, ord_personnel


# Start command
@bot.message_handler(commands=['start'])
def start(message):
    sheet_names = get_sheet_names(service, SAMPLE_SPREADSHEET_ID)
    if sheet_names:
        sheet_list = "\n".join(f"{i + 1}: {name}" for i, name in enumerate(sheet_names))
        bot.reply_to(message, f"Welcome! \n\nPlease select a spreadsheet:\n\n{sheet_list}")
        bot.register_next_step_handler(message, handle_sheet_selection)
    else:
        bot.reply_to(message, "No sheets found in the spreadsheet.")

# Handle sheet selection with error handling for non-integer input
def handle_sheet_selection(message):
    global SHEET_NAME, SAMPLE_RANGE_NAME
    print("Handling sheet selection...")
    bot.reply_to(message, "Processing your selection... Please wait.")  # Notify the user about the processing
    sheet_names = get_sheet_names(service, SAMPLE_SPREADSHEET_ID)

    try:
        # Attempt to convert the user input to an integer
        sheet_index = int(message.text) - 1
        if 0 <= sheet_index < len(sheet_names):
            SHEET_NAME = sheet_names[sheet_index]
            SAMPLE_RANGE_NAME = f"{SHEET_NAME}!{DEFAULT_RANGE}"
            print(f"Sheet '{SHEET_NAME}' selected.")
            bot.reply_to(message, f"Sheet '{SHEET_NAME}' selected. Now, please enter the day of the month (e.g., '9').")
            bot.register_next_step_handler(message, fetch_roster_for_day)
        else:
            print("Invalid sheet selection.")
            bot.reply_to(message, "Invalid selection. Please enter a number corresponding to a sheet.")
            bot.register_next_step_handler(message, handle_sheet_selection)
    except ValueError:
        print("Non-integer input received for sheet selection.")
        bot.reply_to(message, "Please enter a valid number corresponding to a sheet.")
        bot.register_next_step_handler(message, handle_sheet_selection)

# Function to determine month based on sheet name
def get_month_from_sheet_name(sheet_name):
    month_mapping = {
        "Jan": "January", "Feb": "February", "Mar": "March", "Apr": "April",
        "May": "May", "Jun": "June", "Jul": "July", "Aug": "August",
        "Sep": "September", "Sept": "September", "Oct": "October",
        "Nov": "November", "Dec": "December"
    }
    for keyword, month in month_mapping.items():
        if keyword in sheet_name:
            return month
    return "Unknown"

# Function to extract year from sheet name
def get_year_from_sheet_name(sheet_name):
    words = sheet_name.split()
    for word in words:
        if len(word) == 2 and word.isdigit():
            return f"20{word}"
    raise ValueError(f"Invalid sheet name format: '{sheet_name}'. Expected a two-digit year in the sheet name.")

def get_day_of_week(day_of_month, sheet_name):
    month_name = get_month_from_sheet_name(sheet_name)
    year = int(get_year_from_sheet_name(sheet_name))
    date = datetime.strptime(f"{day_of_month} {month_name} {year}", "%d %B %Y")
    return date.strftime("%A").upper()  # Convert day of week to uppercase



# Change this if there are any changes to the MO, SM, and SA
mo_1 = "CPT (DR) CHONG YUAN KAI"
mo_2 = "CPT (DR) NG JIE QI"
sm_1 = "ME3 KARRIE YAP"
sm_2 = "ME2 CHESTON CHEE"
sa_1 = "LCP HOVAN TAN"

medical_officers = [mo_1, mo_2]
senior_medics = [sm_1, sm_2]
assistants = [sa_1]
total_mo_sm_asa = len(medical_officers + senior_medics + assistants)

# Function to fetch and print roster for a specific day
def fetch_roster_for_day(message):
    if message.text.lower() == "back":
        start(message)
        return

    try:
        day_of_month = int(message.text)
       
        bot.reply_to(message, "Fetching data from Google Sheets... Please wait.")


        df = fetch_sheet_data()


        if df.empty:
            bot.reply_to(message, "No data available in the selected sheet. Please check the range or try another sheet.")
            return

        bot.reply_to(message, "Generating Parade State... This may take a few seconds.")

        # Retrieve total_medics and other data from print_roster_for_day
        duty_pairing, duty_personnel, non_duty_personnel, total_medics, ord_personnel = print_roster_for_day(df, day_of_month)


        # Fetch day of the week
        day_of_week = get_day_of_week(day_of_month, SHEET_NAME)

        # Initialize lists for personnel categories
        ojt_personnel = []
        additional = []
        
        # Use ord_personnel for filtering out personnel from duty and non-duty lists
        duty_personnel = duty_personnel[~duty_personnel['Personnel Names'].isin(ord_personnel)]
        non_duty_personnel = non_duty_personnel[~non_duty_personnel['Personnel Names'].isin(ord_personnel)]

        # Filter out "OWADIO" personnel directly in `duty_personnel` and `non_duty_personnel`
        duty_personnel = duty_personnel[~duty_personnel['Personnel Names'].isin(ord_personnel)]
        non_duty_personnel = non_duty_personnel[~non_duty_personnel['Personnel Names'].isin(ord_personnel)]

        # Determine if the day is a weekend
        current_year = datetime.now().year
        current_month = datetime.now().month
        date_to_check = datetime(current_year, current_month, day_of_month)
        is_weekend = date_to_check.weekday() >= 5  # Saturday is 5, Sunday is 6

        # Expanded duty types to include both variations (e.g., "M1" and "M-1")
        duty_types = {'M1': [], 'M2': [], 'M3': [], 'M4': [], 'C1': [], 'C5': []}
        duty_variations = {
            'M1': ["M1", "M-1"], 'M2': ["M2", "M-2"], 'M3': ["M3", "M-3"],
            'M4': ["M4", "M-4"], 'C1': ["C1", "C-1"], 'C5': ["C5", "C-5"]
        }
        prioritized_prefixes = ["AM", "OH"]

        for index, row in duty_personnel.iterrows():
            name = row['Personnel Names']
            duty = row['Duty for Day']

            # Check each duty type and add personnel, prioritizing those with "AM" or "OH" 
            for duty_type, variations in duty_variations.items():
                if any(variation in duty for variation in variations):
                    if any(duty.startswith(prefix) for prefix in prioritized_prefixes):
                        # Insert at the start if "AM" or "OH" precedes the duty
                        duty_types[duty_type].insert(0, name)
                    else:
                        duty_types[duty_type].append(name)

        # Find OJT personnel directly from duty_pairing
        ojt_entries = duty_pairing[duty_pairing['Duty for Day'].str.contains('OJT', na=False)]
        for index, row in ojt_entries.iterrows():
            name = row['Personnel Names']
            duty = row['Duty for Day']
            ojt_personnel.append(f"{name}: {duty}")

        additional = [name for name in additional if name not in [entry.split(":")[0] for entry in ojt_personnel]]

        # Construct duty personnel message with blank line after M4
        duty_personnel_message = "\n".join(
            f"{duty_type}: {' / '.join(personnel_list)}"
            for duty_type, personnel_list in duty_types.items() if personnel_list
        )
        if "M4:" in duty_personnel_message and "C1:" in duty_personnel_message:
            duty_personnel_message = duty_personnel_message.replace(
                f"M4: {' / '.join(duty_types['M4'])}",
                f"M4: {' / '.join(duty_types['M4'])}\n"
            )

        # Calculate additional and other duties
        other_duties = []
        not_present_keywords = ['OFF', 'OSL', 'LL', 'OIL', 'DO', 'MC']
        not_present_count = 0

        personnel_with_owadio = set(df[df.apply(lambda row: row.str.contains("OWADIO", na=False)).any(axis=1)][1])
        for index, row in non_duty_personnel.iterrows():
            duty = row['Duty for Day']
            name = row['Personnel Names']
            if (pd.isna(duty) or duty == '') and name != "CPC" and name not in personnel_with_owadio:
                additional.append(name)
            else:
                if duty and not any(keyword in duty for keyword in keywords_for_duty_personnel):
                    other_duties.append(f"{name}: {duty}")
                    not_present_count += 1

        # Debugging print statements
        print("Duty Personnel:\n", duty_personnel_message)
        print("Additional personnel (excluding 'CPC' and those with 'OWADIO' duties):", additional)
        print("Other Duties (Not Present):", other_duties)

        # First, calculate total_strength, excluding ord_personnel
        total_Strength = total_medics + total_mo_sm_asa - len(ord_personnel)

        # Check if it's a weekend
        is_weekend = day_of_week in ["SATURDAY", "SUNDAY"]

        # Initialize dynamic_medic_strength based on total medics, adjusted for weekends
        if is_weekend:
            # Count occurrences of "CPC" in the M1 and C1 duty lists
            cpc_count_m1 = sum("CPC" in duty for duty in duty_types.get("M1", []))
            cpc_count_c1 = sum("CPC" in duty for duty in duty_types.get("C1", []))
            total_cpc_count = cpc_count_m1 + cpc_count_c1
            
            # Set dynamic_medic_strength to 2 and subtract the CPC occurrences
            dynamic_medic_strength = max(2 - total_cpc_count, 0)
            # Do not add 4 for MO, SM, and SA on weekends
            dynamic_current_strength = dynamic_medic_strength
        else:
            # For weekdays, calculate dynamic_medic_strength normally
            dynamic_medic_strength = total_medics - not_present_count - len(ord_personnel)
            # Add 4 for MO, SM, and SA on weekdays
            dynamic_current_strength = dynamic_medic_strength + 4


        roster_message = f"""
PARADE STATE FOR {day_of_month} {get_month_from_sheet_name(SHEET_NAME)} {get_year_from_sheet_name(SHEET_NAME)}
{day_of_week}
Total Strength: {total_Strength} 
Current Strength: {dynamic_current_strength}/{total_Strength} 
Medic Strength: {dynamic_medic_strength}/{total_medics}

MO: 
{mo_1}:
{mo_2}:

SM:
{sm_1}:
{sm_2}:

Medics:
"""
        if other_duties:
            roster_message += "\n".join(other_duties) + "\n\n"

        roster_message += duty_personnel_message

        # Show additional section only if it's not a weekend
        if additional and not is_weekend:
            additional_section = "\n" + "\nAdditional:\n" + "\n".join(additional)
            roster_message += additional_section
            print("Additional section added to the message.")
        else:
            print("No additional personnel to add or it's a weekend.")

        # OJT section if personnel with "OJT" in duty
        if ojt_personnel:
            ojt_section = "\n\nOJT:\n" + "\n".join(ojt_personnel)
            roster_message += ojt_section
            print("OJT section added to the message.")

        # Final sections
        last_section = f"""

SAR: 
BASE E: 

SUPPLY ASSISTANT:
{sa_1}:

Flying Hours: TBC
"""
        roster_message += last_section
        print(f"Final Message to User:\n{roster_message.strip()}")
        bot.reply_to(message, roster_message.strip())

    except ValueError:
        bot.reply_to(message, "Invalid input. Please enter a numerical day of the month or type 'back' to select a different sheet.")
        bot.register_next_step_handler(message, fetch_roster_for_day)

# Remove any active webhook to avoid conflicts during startup
bot.remove_webhook()

# Run the Flask app with background scheduler thread
if __name__ == "__main__":
    threading.Thread(target=run_scheduler, daemon=True).start()  # Start scheduler in background
    with app.app_context():
        set_webhook()
    
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

