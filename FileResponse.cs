using Aspose.Cells;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Runtime.InteropServices.JavaScript;
using System.Text;

namespace Abi_REST
{
    public class FileResponse
    {
        public async Task<string> ReadFile()
        {
            string path = Path.Combine(Environment.CurrentDirectory, "Files/file.xls");
            var workbook = new Workbook(path);
            workbook.Save(Path.Combine(
                Environment.CurrentDirectory, "Files/out.json"));
            string file = await File.ReadAllTextAsync(Path.Combine(
                Environment.CurrentDirectory, "Files/out.json"));
            return file;
        }
    }
}
