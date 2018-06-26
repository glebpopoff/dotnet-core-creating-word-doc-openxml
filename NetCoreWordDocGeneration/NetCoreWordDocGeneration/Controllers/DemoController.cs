using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using NetCoreWordDocGeneration.Dtos;

namespace NetCoreWordDocGeneration.Controllers
{
    [Route("api/[controller]")]
    [Produces("application/ms-word")]
    public class DemoController : Controller
    {
        [HttpGet("Export")]
        [Produces("application/ms-word")]
        public async Task<IActionResult> Export()
        {
            try
            {
                var demoDto = new DemoDto() { Welcome = "Lorem Ipsum", HelloWorld = "Hello World!!!" };
                return Ok(demoDto);
            }
            catch (Exception ex)
            {
                //log the exception
                return BadRequest();
            }
        }
    }
}
