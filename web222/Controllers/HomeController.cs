using System.Diagnostics;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using web222.Models;

namespace web222.Controllers;

[Authorize]
public class HomeController : Controller
{
    public IActionResult Index()
    {
        return View();
    }

    [Authorize(Roles = "Admin")]
    public IActionResult AdminOnly()
    {
        return View();
    }

    [Authorize(Roles = "User")]
    public IActionResult UserOnly()
    {
        return View();
    }

    [Authorize(Roles = "Guest")]
    public IActionResult GuestOnly()
    {
        return View();
    }
}
