using Microsoft.AspNetCore.Mvc;

namespace web222.Controllers;

public class Controller1 : Controller
{
    // GET
    public IActionResult Index()
    {
        return View();
    }
}