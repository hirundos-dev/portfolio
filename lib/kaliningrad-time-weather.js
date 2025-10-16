      // Функция для обновления времени в Калининграде
      function updateKaliningradTime() {
        const now = new Date();
        const kaliningradTime = new Date(now.toLocaleString("en-US", {timeZone: "Europe/Kaliningrad"}));
        
        const timeElement = document.getElementById('time');
                
        if (timeElement) {
          // Форматируем время (13:00)
          const hours = kaliningradTime.getHours().toString().padStart(2, '0');
          const minutes = kaliningradTime.getMinutes().toString().padStart(2, '0');
          timeElement.textContent = `${hours}:${minutes}`;
        }
      }
      
      // Функция для получения погоды в Калининграде
      async function updateWeather() {
        try {
          // Используем бесплатный API для погоды
          const response = await fetch('https://api.open-meteo.com/v1/forecast?latitude=54.7074&longitude=20.5073&current=temperature_2m,weather_code&timezone=Europe/Kaliningrad');
          const data = await response.json();
          
          const weatherDescElement = document.getElementById('weather-desc');
          const weatherTempElement = document.getElementById('weather-temp');
          
          if (weatherDescElement && weatherTempElement) {
            const temp = Math.round(data.current.temperature_2m);
            const weatherCode = data.current.weather_code;
            
            // Преобразуем код погоды в описание
            const weatherDescriptions = {
              0: 'ясно',
              1: 'малооблачно',
              2: 'облачно',
              3: 'пасмурно',
              45: 'туман',
              48: 'туман',
              51: 'морось',
              53: 'морось',
              55: 'морось',
              61: 'дождь',
              63: 'дождь',
              65: 'дождь',
              71: 'снег',
              73: 'снег',
              75: 'снег',
              95: 'гроза'
            };
            
            const description = weatherDescriptions[weatherCode] || 'Облачно';
            weatherDescElement.textContent = description;
            weatherTempElement.textContent = `${temp}°C`;
          }
        } catch (error) {
          console.log('Ошибка получения погоды:', error);
          // Fallback значения
          const weatherDescElement = document.getElementById('weather-desc');
          const weatherTempElement = document.getElementById('weather-temp');
          if (weatherDescElement && weatherTempElement) {
            weatherDescElement.textContent = 'Облачно';
            weatherTempElement.textContent = '+15°C';
          }
        }
      }
      
      // Обновляем время сразу и каждую минуту
      updateKaliningradTime();
      setInterval(updateKaliningradTime, 60000); // Обновляем каждую минуту
      
      // Обновляем погоду сразу и каждые 30 минут
      updateWeather();
      setInterval(updateWeather, 1800000); // Обновляем каждые 30 минут