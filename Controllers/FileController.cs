using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;

namespace Abi_REST
{
    [Route("api/[controller]")]
    [ApiController]
    public class AppController:ControllerBase
    {
        [HttpGet]
        public async Task<IActionResult> GetFile()
        {
            var file = new FileResponse();
            var json=await file.ReadFile();
            return Ok(json);
        }
    }
}
