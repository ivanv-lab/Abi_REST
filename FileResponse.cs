using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace Abi_REST
{
    public class FileResponse
    {
        public string Name {  get; set; }
        public DateOnly startDate {  get; set; }
        public DateOnly endDate { get; set; }
        public string Url {  get; set; }

        public async Task<JObject> ReadFile()
        {
            //string path = "C:\\Users\\adminvl\\Desktop\\File\\file.xls";

            Microsoft.Office.Interop.Excel.Application
                ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook =
                ObjExcel.Workbooks.Open("C:\\Users\\adminvl\\Desktop\\File\\file.xls", 0, false, 5, "", "", false,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet worksheet =
                (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            List<string> strings = new List<string>();

            for (int i = 0; i < 4; i++)
            {
                Microsoft.Office.Interop.Excel.Range usedColumn =
                    (Microsoft.Office.Interop.Excel.Range)
                    worksheet.UsedRange.Columns[i];
                JObject json = (JObject)usedColumn.Cells.Value2;
                strings.Add(json.ToString());
            }

            return new JObject(strings.ToArray());

            //string path = "C:\\Users\\adminvl\\Desktop\\File\\file.xls";
            //var file = await File.ReadAllLinesAsync(path);
            //JObject json = null;
            //foreach (var line in file)
            //{
            //    json = JObject.FromObject(line);
            //}
            //return json;
        }
    }
}
