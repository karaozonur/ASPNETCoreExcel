using ASPNETCoreExcel.Models;
using ClosedXML.Excel;

namespace ASPNETCoreExcel.Services
{
    public class ExcelImportService
    {
        public List<Person> ImportFromExcel(Stream fileStream)
        {
            var personList = new List<Person>();

            using (var workbook = new XLWorkbook(fileStream))
            {
                var worksheet = workbook.Worksheets.First();
                var rows = worksheet.RowsUsed();

                foreach (var row in rows.Skip(1)) // İlk satır başlık olduğu için atlıyoruz
                {
                    var person = new Person
                    {
                        Id = int.Parse(row.Cell(1).GetString()),
                        Name = row.Cell(2).GetString(),
                        Email = row.Cell(3).GetString()
                    };

                    personList.Add(person);
                }
            }

            return personList;
        }
    }
}
