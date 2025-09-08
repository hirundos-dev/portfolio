      // Функция для обновления времени в Калининграде
      function updateKaliningradTime() {
        const now = new Date();
        const kaliningradTime = new Date(now.toLocaleString("en-US", {timeZone: "Europe/Kaliningrad"}));
        
        const timeElement = document.querySelector('.footer__time .time');
        const dateElement = document.querySelector('.footer__time .date');
        
        if (timeElement && dateElement) {
          // Форматируем время (13:00)
          const hours = kaliningradTime.getHours().toString().padStart(2, '0');
          const minutes = kaliningradTime.getMinutes().toString().padStart(2, '0');
          timeElement.textContent = `${hours}:${minutes}`;
          
          // Форматируем дату (12.12.1990)
          const day = kaliningradTime.getDate().toString().padStart(2, '0');
          const month = (kaliningradTime.getMonth() + 1).toString().padStart(2, '0');
          const year = kaliningradTime.getFullYear();
          dateElement.textContent = `${day}.${month}.${year}`;
        }
      }
      
      // Функция для получения погоды в Калининграде
      async function updateWeather() {
        try {
          // Используем бесплатный API для погоды
          const response = await fetch('https://api.open-meteo.com/v1/forecast?latitude=54.7074&longitude=20.5073&current=temperature_2m,weather_code&timezone=Europe/Moscow');
          const data = await response.json();
          
          const weatherDescElement = document.querySelector('.footer__weather .weather-desc');
          const weatherTempElement = document.querySelector('.footer__weather .weather-temp');
          
          if (weatherDescElement && weatherTempElement) {
            const temp = Math.round(data.current.temperature_2m);
            const weatherCode = data.current.weather_code;
            
            // Преобразуем код погоды в описание
            const weatherDescriptions = {
              0: 'Ясно',
              1: 'Малооблачно',
              2: 'Облачно',
              3: 'Пасмурно',
              45: 'Туман',
              48: 'Туман',
              51: 'Морось',
              53: 'Морось',
              55: 'Морось',
              61: 'Дождь',
              63: 'Дождь',
              65: 'Дождь',
              71: 'Снег',
              73: 'Снег',
              75: 'Снег',
              95: 'Гроза'
            };
            
            const description = weatherDescriptions[weatherCode] || 'Облачно';
            weatherDescElement.textContent = description;
            weatherTempElement.textContent = `${temp}°C`;
          }
        } catch (error) {
          console.log('Ошибка получения погоды:', error);
          // Fallback значения
          const weatherDescElement = document.querySelector('.footer__weather .weather-desc');
          const weatherTempElement = document.querySelector('.footer__weather .weather-temp');
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