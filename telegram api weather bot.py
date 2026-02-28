# for tg bot
import requests
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# for meteo
import openmeteo_requests
import pandas as pd
import requests_cache
from retry_requests import retry
cache_session = requests_cache.CachedSession('.cache', expire_after = 3600)
retry_session = retry(cache_session, retries = 5, backoff_factor = 0.2)
openmeteo = openmeteo_requests.Client(session = retry_session)

def get_forecast(lat,lon):
    url = "https://api.open-meteo.com/v1/forecast"
    params = {
        "latitude": lat,
        "longitude": lon,
        "hourly": ["temperature_2m", "rain", "snowfall", "wind_speed_10m", "cloud_cover", "apparent_temperature"],
        "current": ["rain", "temperature_2m", "apparent_temperature", "snowfall"],
        "minutely_15": ["temperature_2m", "rain", "snowfall", "wind_speed_10m"],
        "timezone": "Europe/Moscow",
        "forecast_days": 3,
        "wind_speed_unit": "ms",
    }
    responses = openmeteo.weather_api(url, params=params)

    # Process first location. Add a for-loop for multiple locations or weather models
    response = responses[0]

    # Process minutely_15 data. The order of variables needs to be the same as requested.
    minutely_15 = response.Minutely15()
    minutely_15_temperature_2m = minutely_15.Variables(0).ValuesAsNumpy()
    minutely_15_rain = minutely_15.Variables(1).ValuesAsNumpy()
    minutely_15_snowfall = minutely_15.Variables(2).ValuesAsNumpy()
    minutely_15_wind_speed_10m = minutely_15.Variables(3).ValuesAsNumpy()

    minutely_15_data = {"date": pd.date_range(
        start = pd.to_datetime(minutely_15.Time() + response.UtcOffsetSeconds(), unit = "s", utc = True),
        end =  pd.to_datetime(minutely_15.TimeEnd() + response.UtcOffsetSeconds(), unit = "s", utc = True),
        freq = pd.Timedelta(seconds = minutely_15.Interval()),
        inclusive = "left"
    )}

    minutely_15_data["temperature_2m"] = minutely_15_temperature_2m
    minutely_15_data["rain"] = minutely_15_rain
    minutely_15_data["snowfall"] = minutely_15_snowfall
    minutely_15_data["wind_speed_10m"] = minutely_15_wind_speed_10m

    minutely_15_dataframe = pd.DataFrame(data = minutely_15_data)

    # Process hourly data. The order of variables needs to be the same as requested.
    hourly = response.Hourly()
    hourly_temperature_2m = hourly.Variables(0).ValuesAsNumpy()
    hourly_rain = hourly.Variables(1).ValuesAsNumpy()
    hourly_snowfall = hourly.Variables(2).ValuesAsNumpy()
    hourly_wind_speed_10m = hourly.Variables(3).ValuesAsNumpy()
    hourly_cloud_cover = hourly.Variables(4).ValuesAsNumpy()
    hourly_apparent_temperature = hourly.Variables(5).ValuesAsNumpy()

    hourly_data = {"date": pd.date_range(
        start = pd.to_datetime(hourly.Time() + response.UtcOffsetSeconds(), unit = "s", utc = True),
        end =  pd.to_datetime(hourly.TimeEnd() + response.UtcOffsetSeconds(), unit = "s", utc = True),
        freq = pd.Timedelta(seconds = hourly.Interval()),
        inclusive = "left"
    )}

    hourly_data["temperature_2m"] = hourly_temperature_2m
    hourly_data["rain"] = hourly_rain
    hourly_data["snowfall"] = hourly_snowfall
    hourly_data["wind_speed_10m"] = hourly_wind_speed_10m
    hourly_data["cloud_cover"] = hourly_cloud_cover
    hourly_data["apparent_temperature"] = hourly_apparent_temperature

    hourly_dataframe = pd.DataFrame(data = hourly_data)

    # use hourly_dataframe or minutely_15_dataframe
    return minutely_15_dataframe


# bot parameters
token = ''
tg_bot_url = ''


# Commands
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Hello!")

    
def handle_resp(text: str)-> str:
    if 'weather' in text:
        return 'wwww'
    return '???'


async def forecast(update: Update, context: ContextTypes.DEFAULT_TYPE):
       for i in range(len(loc)//2):
            df = get_forecast(loc[2*i], loc[2*i+1])
            print('ok')
            answer:str = ''
            for j in range(28,41):
                date = str(df.loc[j,'date'])[5:16]
                temp = float(df.loc[j,'temperature_2m'])
                rain = float(df.loc[j,'rain']) + float(df.loc[j,'snowfall'])
                answer = answer+str(date)+str(temp)+str(rain)+'\n'
            answer = answer +'\n'
            for j in range(68,81):
                date = str(df.loc[j,'date'])[5:16]
                temp = round(float(df.loc[j,'temperature_2m']),1)
                rain = round(float(df.loc[j,'rain']) + float(df.loc[j,'snowfall']),1)
                answer = answer+str(date)+' ' + str(temp)+' '+str(rain)+'\n'


            await update.message.reply_text(f'location {i+1} \n {answer}')



async def locations(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if loc == []:
        await update.message.reply_text('enter location latitude and longitude')
    else:
        await update.message.reply_text(f'locations are: {loc}')

# Message handler
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message_type: str = update.message.chat.type
    text: str = update.message.text
    print(text)
    numbers = text.split()
    lat = numbers[0]
    lon = numbers[1]
    await update.message.reply_text(f'all good {lat}, man {lon}')
    loc.append(lat)
    loc.append(lon)


async def error(update: Update, context: ContextTypes.DEFAULT_TYPE):
    print(f'Update {update} caused error {context.error}')

if __name__ == '__main__':
    app = Application.builder().token(token).build()

    loc = []
    # app.add_handler()
    app.add_handler(CommandHandler('start',start_command))
    app.add_handler(CommandHandler('forecast',forecast))
    app.add_handler(CommandHandler('locations',locations))

    app.add_handler(MessageHandler(filters.TEXT, handle_message))

    app.add_error_handler(error)
    print('ammmmmm')

    app.run_polling(poll_interval=5)
