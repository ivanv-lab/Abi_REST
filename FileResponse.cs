using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;

namespace Abi_REST
{
    public class FileResponse
    {
        public string Name {  get; set; }
        public string startDate {  get; set; }
        public string endDate { get; set; }
        public string Url {  get; set; }

        public async Task<List<string>> ReadFile()
        {
            string path = Path.Combine(Environment.CurrentDirectory, "file.xls");

            //Ошибка библиотеки
            //using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, false))
            //{
            //    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart
            //        ?? spreadsheetDocument.AddWorkbookPart();
            //    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

            //    OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
            //    List<string> rows = new List<string>();
            //    while (reader.Read())
            //    {
            //        if (reader.ElementType == typeof(FileResponse))
            //        {
            //            rows.Add(reader.GetText());
            //        }
            //    }
            //    return rows;
            //}

            //Какие то приколы с кодировкой
            //var list = new List<FileResponse>();
            //using (var stream = File.Open(path, FileMode.Open,
            //    FileAccess.Read))
            //{
            //    using (var reader = ExcelReaderFactory
            //        .CreateReader(stream))
            //    {
            //        while (reader.Read())
            //        {
            //            var obj = new FileResponse
            //            {
            //                Name = reader.GetString(0),
            //                startDate = reader.GetString(1),
            //                endDate = reader.GetString(2),
            //                Url = reader.GetString(3)
            //            };
            //            list.Add(obj);
            //        }
            //        reader.Close();
            //    }
            //}
            //var json = JsonConvert.SerializeObject(list);
            //return json;

            //Способ работал бы, если бы можно было назвать диапазон
            //таблицей
            //string sheetName = "sheet1";
            //string newPath = Path.Combine(
            //    Environment.CurrentDirectory, "out.json");
            //string json;

            //var connectionString = $@"
            //    Provider=Microsoft.ACE.OLEDB.12.0;
            //    Data Source={path};
            //    Extended Properties=""Excel 12.0 Xml;HDR=YES""
            //        ";
            //using (var conn = new OleDbConnection(connectionString))
            //{
            //    conn.Open();

            //    var cmd = conn.CreateCommand();
            //    cmd.CommandText = $"select * from [Название_тендера]";

            //    using (var rdr = cmd.ExecuteReader())
            //    {
            //        var query = rdr.Cast<List<FileResponse>>()
            //            .Select(row => new
            //            {
            //                Name = row[0],
            //                startDate = row[1],
            //                endDate = row[3],
            //                Url = row[4]
            //            });
            //        json = JsonConvert.SerializeObject(query);
            //    }
            //}
            //return json;

            //По идее может работать, но пока проблемы с парсингом
            //НУЖНО ЧТО ТО ПРИДУМАТЬ С ЧТЕНИЕМ ДОКУМЕНТА ОБРАТНО
            //ИЛИ КАК ТО ДОБАВИТЬ КЛЮЧИ Item
            //var workbook=new Workbook(path);
            //workbook.Save(Path.Combine(
            //    Environment.CurrentDirectory,"out.json"));
            //var file =await File.ReadAllTextAsync(Path.Combine(
            //    Environment.CurrentDirectory, "out.json"));
            //JObject json = JObject.Parse(file);
            //foreach (var item in json[""])
            //{
            //    Name = item["Название тендера"].ToString();
            //}
            //return json;

            //Не рабочий
            //byte[] file = await File.ReadAllBytesAsync(path);
            //string lines = Encoding.UTF8.GetString(file);
            //JObject json = null;
            //foreach (var line in lines)
            //{
            //    json = JObject.Parse(line.ToString());
            //}
            //return json;
        }
    }
}
