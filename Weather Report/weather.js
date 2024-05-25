let input = document.querySelector("input");
let btn = document.querySelector("button");
let CityName = document.querySelectorAll("h4")[0];
let temp = document.querySelectorAll("h4")[1];
let WindSpeed = document.querySelectorAll("h4")[2];
console.log(WindSpeed);
let Humidity = document.querySelectorAll("h4")[3];
console.log(Humidity);
let Climate = document.querySelectorAll("h4")[4];

btn.addEventListener("click", async () =>
{
    let apiKey = `218697af7b01ae22a137771634403c9d`;

    let dataServer = await fetch(
      `https://api.openweathermap.org/data/2.5/weather?q=${input.value}&appid=${apiKey}`,
    );
    let data = await dataServer.json();
    console.log(data);
    CityName.innerHTML = `CityName  :  ${data.name}`;

    temp.innerHTML = `Temperature  : ${(data.main.temp-273).toFixed(2)} °C`;

    WindSpeed.innerHTML = `Wind Speed  : ${data.wind.speed}`;

    Humidity.innerHTML = `Humidity : ${data.main.humidity}`;

    Climate.innerHTML  = `Climate : ${data.weather[0].main}`;
});