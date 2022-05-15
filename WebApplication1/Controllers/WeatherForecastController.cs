using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public IEnumerable<WeatherForecast> Get()
        {
            var rng = new Random();
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = rng.Next(-20, 55),
                Summary = Summaries[rng.Next(Summaries.Length)]
            })
            .ToArray();
        }
        [HttpGet("download")]
        public IActionResult download()
        {
            var path = @"C:\Users\Ayad\source\repos\ClosedXMLPOC\WebApplication1\DynamicImageProvider.png";
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Inserting Tables");
            ws.Row(1).Height = 100;
            var image = ws.AddPicture(path)
                .MoveTo(ws.Cell(1,1));
            //image.Scale(.25);


            // From a query
            var list = new List<Person>();
            list.Add(new Person() { Name = "John", Age = 30, House = "On Elm St." });
            list.Add(new Person() { Name = "Mary", Age = 15, House = "On Main St." });
            list.Add(new Person() { Name = "Luis", Age = 21, House = "On 23rd St." });
            list.Add(new Person() { Name = "Henry", Age = 45, House = "On 5th Ave." });

            var people = from p in list
                         where p.Age >= 21
                         select new { p.Name, p.House, p.Age };

            ws.Cell(3, 1).Value = "From Query";
            //ws.Range(3, 6, 7, 8).Merge().AddToNamed("Titles");
            var tableWithPeople = ws.Cell(4, 1).InsertTable(people.AsEnumerable());
            // Format all titles in one shot
            //wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;

            ws.Columns().AdjustToContents();

            wb.SaveAs("InsertingTables.xlsx");
            byte[] data;
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                data = stream.ToArray();
            }
            return File(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "sdsd");
        }
    }
    class Person
    {
        public String House { get; set; }
        public String Name { get; set; }
        public Int32 Age { get; set; }
    }
}
