import discord, xlwt
from discord.ext import commands
from datetime import datetime

## THIS IS A WORK IN PROGRESS! NOT QUITE FUNCTIONAL YET, BUT HOPEFULLY SOON!
## Bot must be able to read messages, message history, and optionally send messages in order to send a DM with the .xlsx.

GUILD_ID = 1234 # ID of the server you want to query
TOKEN = "" # Discord bot token

intents = discord.Intents().all()
client = commands.Bot(command_prefix = ",", intents = intents)

@client.event
async def on_ready():
    print("[!] Excel bot is ready!\n")

    data, dates = await getData()
    exportExcel(data, dates, "data.xlsx")
    #uploadExcel()

async def getData():
    data = []
    allDates = []
    server = client.get_guild(GUILD_ID)

    for member in server.members:
        if not member.bot:
            data.append({ "username": member.name, "profilePictureURL": str(member.avatar_url), "dates": [] })

    for channel in server.text_channels:
        try:
            messages = await channel.history().flatten()
            for message in messages:
                for i in range(len(data)):
                    if data[i]["username"] == message.author.name:
                        if len(data[i]["dates"]) > 0:
                            if data[i]["dates"][-1]["date"] == str(message.created_at.year) + "-" + str(message.created_at.month) + "-" + str(message.created_at.day):
                                data[i]["dates"][-1]["count"] += 1
                                if str(message.created_at.year) + "-" + str(message.created_at.month) + "-" + str(message.created_at.day) not in allDates:
                                    allDates.append(str(message.created_at.year) + "-" + str(message.created_at.month) + "-" + str(message.created_at.day))
                            else:
                                data[i]["dates"].append({ "date": str(message.created_at.year) + "-" + str(message.created_at.month) + "-" + str(message.created_at.day), "count": data[i]["dates"][-1]["count"] + 1 })
                                if str(message.created_at.year) + "-" + str(message.created_at.month) + "-" + str(message.created_at.day) not in allDates:
                                    allDates.append(str(message.created_at.year) + "-" + str(message.created_at.month) + "-" + str(message.created_at.day))
                        else:
                            data[i]["dates"].append({ "date": str(message.created_at.year) + "-" + str(message.created_at.month) + "-" + str(message.created_at.day), "count": 1 })
                            if str(message.created_at.year) + "-" + str(message.created_at.month) + "-" + str(message.created_at.day) not in allDates:
                                allDates.append(str(message.created_at.year) + "-" + str(message.created_at.month) + "-" + str(message.created_at.day))  
        except discord.errors.Forbidden:
            pass

    for i in range(len(data)):
        cleanedDates = []

        for y in range(len(data[i]["dates"])):
            for entry in cleanedDates:
                if entry["date"] == data[i]["dates"][y]["date"]:
                    break
                cleanedDates.append({ "date": data[i]["dates"][y]["date"], "count": data[i]["dates"][y]["count"] })

            # Pick up here:
            # TO DO: Adjust sorting.

        data[i]["dates"] = cleanedDates
        data[i]["dates"].sort(key = lambda x: datetime.strptime(x["date"], '%Y-%m-%d'), reverse = True)

    allDates.sort(key = lambda x: datetime.strptime(x, '%Y-%m-%d'), reverse = True)

    temp = []

    for i in range(len(data[5]["dates"])):
        temp.append(data[5]["dates"][i]["date"])

    print(allDates)
    print(temp)

    return data, allDates

def exportExcel(data, dates, name = "data.xlsx"):
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet("Data", cell_overwrite_ok = True)
    #worksheet = workbook.add_sheet("Data")

    worksheet.write(0, 0, "Discord Users")
    worksheet.write(0, 1, "")
    worksheet.write(0, 2, "")

    currentRow = 2
    currentColumn = 3
    createdColumnDates = []
    createdColumnPositions = []
    
    for i in range(len(data)):
        worksheet.write(currentRow, 0, data[i]["username"])
        worksheet.write(currentRow, 2, data[i]["profilePictureURL"])

        for y in range(len(data[i]["dates"])):
            if data[i]["dates"][y]["date"] not in createdColumnDates:
                worksheet.write(0, currentColumn, data[i]["dates"][y]["date"])
                worksheet.write(currentRow, currentColumn, data[i]["dates"][y]["count"])
                createdColumnDates.append(data[i]["dates"][y]["date"])
                createdColumnPositions.append(currentColumn)

                currentColumn += 1
            else:
                for z in range(len(createdColumnDates)):
                    if data[i]["dates"][y]["date"] == createdColumnDates[z]:
                         worksheet.write(currentRow, createdColumnPositions[z], data[i]["dates"][y]["count"])
        currentRow += 1

    workbook.save(name)
    print("Saved excel file.")

client.run(TOKEN)
