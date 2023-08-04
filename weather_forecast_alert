function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Weather Forecast')
    .addItem('Get Weather Forecast', 'getWeatherForecast')
    .addToUi();
}

function getWeatherForecast() {
  var apiEndpoint = "https://api.openweathermap.org/data/2.5/forecast";
  var latitude = 13.732527;
  var longitude = 100.516766;
  var api_key = "...";
  var units = "metric";
  var cnt = 6;

  var url = `${apiEndpoint}?lat=${latitude}&lon=${longitude}&appid=${api_key}&units=${units}&cnt=${cnt}`;

  try {
    var response = UrlFetchApp.fetch(url);
    var forecast_data = JSON.parse(response.getContentText());
    if (forecast_data) {
      var weatherDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WeatherData");
      weatherDataSheet.clearContents();

      var weather_emoji = {
        200: "⛈️ thunderstorm with light rain",
        201: "⛈️ thunderstorm with rain",
        202: "⛈️ thunderstorm with heavy rain",
        210: "⛈️ light thunderstorm",
        211: "⛈️ thunderstorm",
        212: "⛈️ heavy thunderstorm",
        221: "⛈️ ragged thunderstorm",
        230: "⛈️ thunderstorm with light drizzle",
        231: "⛈️ thunderstorm with drizzle",
        232: "⛈️ thunderstorm with heavy drizzle",
        300: "🌧️ light intensity drizzle",
        301: "🌧️ drizzle",
        302: "🌧️ heavy intensity drizzle",
        310: "🌧️ light intensity drizzle rain",
        311: "🌧️ drizzle rain",
        312: "🌧️ heavy intensity drizzle rain",
        313: "🌧️ shower rain and drizzle",
        314: "🌧️ heavy shower rain and drizzle",
        321: "🌧️ shower drizzle",
        500: "🌧️ light rain",
        501: "🌧️ moderate rain",
        502: "🌧️ heavy intensity rain",
        503: "🌧️ very heavy rain",
        504: "🌧️ extreme rain",
        511: "🌧️ freezing rain",
        520: "🌧️ light intensity shower rain",
        521: "🌧️ shower rain",
        522: "🌧️ heavy intensity shower rain",
        531: "🌧️ ragged shower rain",
      };

      var dataToAppend = [];
      for (var i = 0; i < forecast_data.list.length; i++) {
        var data = forecast_data.list[i];
        var pop = data.pop * 100;
        var weather_id = data.weather[0].id;
        if (200 <= weather_id && weather_id < 600) {
          var weather_description = data.weather[0].description;
          var emoji_description = weather_emoji[weather_id] || weather_description;
          var row = [data.dt_txt, pop, emoji_description];
          dataToAppend.push(row);
        }
      }

      if (dataToAppend.length > 0) {
        weatherDataSheet.getRange(1, 1, dataToAppend.length, dataToAppend[0].length).setValues(dataToAppend);
        sendLineMessage(dataToAppend);
      } else {
        Logger.log("No relevant weather data found.");
      }
    } else {
      Logger.log("Failed to fetch weather forecast data.");
    }
  } catch (e) {
    Logger.log("Error occurred: " + e);
  }
}

function sendLineMessage(weatherData) {
  var channelAccessToken = "...";
  var lineEndpoint = "https://api.line.me/v2/bot/message/broadcast";
  var headers = {
    "Authorization": "Bearer " + channelAccessToken,
    "Content-Type": "application/json",
  };

  var message = "Weather Forecast:\n";
  var date = weatherData[0][0].split(" ")[0];
  message += `${date}\n`;

  for (var i = 0; i < weatherData.length; i++) {
    var data = weatherData[i];
    var timeUTC = new Date(data[0]);
    var time = Utilities.formatDate(timeUTC, "GMT+7", "HH:mm");
    var pop = data[1];
    var weather = data[2];

    message += `${time} POP. ${pop}% ${weather}\n`;
  }

  message += "☂️ Bring an Umbrella ☂️";

  var payload = {
    "messages": [
      {
        "type": "text",
        "text": message,
      }
    ],
  };

  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true,
  };

  var response = UrlFetchApp.fetch(lineEndpoint, options);
  Logger.log("LINE API Response: " + response.getContentText());
}
