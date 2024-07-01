using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using web222.Context;
using web222.Models;

namespace web222.Controllers
{
    [Authorize(Roles = "Admin")]
    public class PracodawcaController : Controller
    {
        private readonly ApplicationDbContext _context;

        public PracodawcaController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Create(Pracodawca pracodawca)
        {
            if (ModelState.IsValid)
            {
                _context.Add(pracodawca);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Create));
            }
            return View(pracodawca);
        }

        [HttpGet]
        public async Task<IActionResult> List()
        {
            var pracodawcy = await _context.Pracodawcy.ToListAsync();
            return View(pracodawcy);
        }
    }
}