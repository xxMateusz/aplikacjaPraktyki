using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Threading.Tasks;
using web222.Context;
using web222.Models;

namespace web222.Controllers
{
    [Authorize(Roles = "Admin")]
    public class ZoneController : Controller
    {
        private readonly ApplicationDbContext _context;

        public ZoneController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public async Task<IActionResult> Index()
        {
            var zones = await _context.Zones.ToListAsync();
            return View(zones);
        }

        [HttpPost]
        public async Task<IActionResult> Index(int[] selectedZones)
        {
            var zones = await _context.Zones.ToListAsync();
            foreach (var zone in zones)
            {
                zone.IsSelected = selectedZones.Contains(zone.Id);
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }
    }
}