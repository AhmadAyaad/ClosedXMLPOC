using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
namespace ClosedXMLPOC
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Inserting Tables");
            // From a query
            var list = new List<Person>();
            list.Add(new Person() { Name = "John", Age = 30, House = "On Elm St." });
            list.Add(new Person() { Name = "Mary", Age = 15, House = "On Main St." });
            list.Add(new Person() { Name = "Luis", Age = 21, House = "On 23rd St." });
            list.Add(new Person() { Name = "Henry", Age = 45, House = "On 5th Ave." });

            var people = from p in list
                         where p.Age >= 21
                         select new { p.Name, p.House, p.Age };

            ws.Cell(7, 6).Value = "From Query";
            ws.Range(7, 6, 7, 8).Merge().AddToNamed("Titles");
            var tableWithPeople = ws.Cell(8, 6).InsertTable(people.AsEnumerable());
            // Format all titles in one shot
            //wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;

            ws.Columns().AdjustToContents();

            wb.SaveAs("InsertingTables.xlsx");
        }
    }
    class Person
    {
        public String House { get; set; }
        public String Name { get; set; }
        public Int32 Age { get; set; }
    }

    //private DataTable GetTable()
    //{
    //    DataTable table = new DataTable();
    //    table..Columns.Add("Dosage", typeof(int));
    //    table.Columns.Add("Drug", typeof(string));
    //    table.Columns.Add("Patient", typeof(string));
    //    table.Columns.Add("Date", typeof(DateTime));

    //    table.Rows.Add(25, "Indocin", "David", DateTime.Now);
    //    table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
    //    table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
    //    table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
    //    table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);
    //    return table;
    //}
}
