import discord
import pandas as pd
from datetime import datetime
import os

# Bot setup
TOKEN = 'MTMwMTYyNTk0NTQ1NDgwNTA2Mw.G4myGw.9djQLcIKulpZk_KL1UN39eDpK7RqQKjKihpNyI'
CHANNEL_ID = 1301619788929564692  # Replace with your target channel ID

client = discord.Client(intents=discord.Intents.default())
messages_data = []


@client.event
async def on_ready():
    print(f'Logged in as {client.user}')


@client.event
async def on_message(message):
    if message.channel.id == CHANNEL_ID and not message.author.bot:
        # Log message data
        messages_data.append({
            "Username": message.author.name,
            "Content": message.content,
            "Timestamp": message.created_at.strftime('%Y-%m-%d %H:%M:%S')
        })
        print(f"Logged message from {message.author.name}")


# Save messages to Excel and clear daily logs
def save_to_excel():
    global messages_data
    if messages_data:
        df = pd.DataFrame(messages_data)
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f'discord_chat_{today}.xlsx'

        df.to_excel(filename, index=False)
        print(f"Saved messages to {filename}")

        # Clear the data for the next day
        messages_data = []


# Run daily export
async def daily_export():
    await client.wait_until_ready()
    while not client.is_closed():
        now = datetime.now()
        if now.hour == 0 and now.minute == 0:  # Set to midnight
            save_to_excel()
            await asyncio.sleep(86400)  # Wait 24 hours for the next export


# Schedule the daily export
import asyncio

client.loop.create_task(daily_export())

client.run(TOKEN)
