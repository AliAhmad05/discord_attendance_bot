import discord
import gspread
import datetime
from oauth2client.service_account import ServiceAccountCredentials

# Discord bot setup
intents = discord.Intents.default()
intents.message_content = True
client = discord.Client(intents=intents)

# Specify the ID of the attendance channel
attendance_channel_id = '948087217367171123'

# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
gc = gspread.authorize(credentials)
spreadsheet_key = '1xq4dM-6KEBhmoYRm0xJhOLIK6t8FfIypOrmvcO4puiQ'  # Replace with your actual spreadsheet key
current_month = datetime.datetime.utcnow().strftime('%B')  # Full month name
worksheet_name = current_month  # Convert to lowercase for consistency

@client.event
async def on_ready():
    print(f'Logged in as {client.user.name}')

@client.event
async def on_message(message):
    # Check if the message author is not a bot and the message content indicates attendance
    if not message.author.bot and 'here' in message.content.lower():
        current_date = datetime.datetime.utcnow().strftime('%-d/%-m')  # Format date as "d/m"
        name = message.author.display_name.lower()  # Convert the name to lowercase for case-insensitive matching

        try:
            worksheet = gc.open_by_key(spreadsheet_key).worksheet(worksheet_name)
            headers = worksheet.row_values(2)  # Get headers from the second row
            values = worksheet.col_values(1)  # Get all values from the first column (names)

            date_column_index = -1
            for index, header in enumerate(headers):
                if current_date in header:
                    date_column_index = index + 1
                    break

            if date_column_index != -1:
                name_matched = False
                for index, row_name in enumerate(values[1:], start=1):
                    if name in row_name.lower():
                        # Check if the cell is empty before updating attendance
                        if not worksheet.cell(index + 1, date_column_index).value:
                            # Update the attendance cell for the matching name and date column
                            worksheet.update_cell(index + 1, date_column_index, 'P')
                            await message.author.send(f"Your attendance for {current_date} has been marked.")
                            await message.add_reaction('✅')  # React to the message with a checkmark emoji
                        else:
                            await message.author.send(f"Your attendance for {current_date} is already marked.")
                        name_matched = True
                        break

                if not name_matched:
                    support_user = 'Ali Ahmad'
                    await message.author.send(f"Sorry, your name does not match any records. Please contact {support_user} for support.")
        
        except gspread.exceptions.WorksheetNotFound:
            await message.channel.send("The specified worksheet was not found. Please contact Ali Ahmad for more information.")

        except Exception as e:
            print(f"Error marking attendance: {e}")

# Replace 'YOUR_BOT_TOKEN' with your actual bot token
client.run('MTE0MDIzMDcyMzkzNzIzOTA1MA.Gv6i39.x1-gjcUw4t90eZdGOXicsG1Un8muZ7gRJDQvDg')
