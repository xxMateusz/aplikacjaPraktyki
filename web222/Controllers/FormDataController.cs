using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Linq;
using System.Threading.Tasks;
using web222.Context;
using web222.Models;

namespace web222.Controllers
{
  //  [Authorize(Roles = "Admin")]
    public class FormDataController : Controller
    {
        private readonly ApplicationDbContext _context;

        public FormDataController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public async Task<IActionResult> Create()
        {
            ViewBag.Pracodawcy = await _context.Pracodawcy.ToListAsync();
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Create(FormData formData)
        {
            if (ModelState.IsValid)
            {
                _context.Add(formData);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Create));
            }
            ViewBag.Pracodawcy = await _context.Pracodawcy.ToListAsync();
            return View(formData);
        }
    }
}