using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Threading.Tasks;
using web222.Context;
using web222.Models;

namespace web222.Controllers
{
    [Authorize(Roles = "Admin")]
    public class AdminController : Controller
    {
        private readonly ApplicationDbContext _context;

        public AdminController(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> Index()
        {
            var formData = await _context.FormData.ToListAsync();
            var pracodawcy = await _context.Pracodawcy.ToListAsync();
            var zones = await _context.Zones.ToListAsync();

            var model = new AdminViewModel
            {
                FormData = formData,
                Pracodawcy = pracodawcy,
                Zones = zones
            };

            return View(model);
        }

        [HttpPost]
        public async Task<IActionResult> DeleteFormData(int id)
        {
            var formData = await _context.FormData.FindAsync(id);
            if (formData != null)
            {
                _context.FormData.Remove(formData);
                await _context.SaveChangesAsync();
            }
            return RedirectToAction(nameof(Index));
        }

        [HttpPost]
        public async Task<IActionResult> DeletePracodawca(int id)
        {
            var pracodawca = await _context.Pracodawcy.FindAsync(id);
            if (pracodawca != null)
            {
                _context.Pracodawcy.Remove(pracodawca);
                await _context.SaveChangesAsync();
            }
            return RedirectToAction(nameof(Index));
        }

        [HttpPost]
        public async Task<IActionResult> DeleteZone(int id)
        {
            var zone = await _context.Zones.FindAsync(id);
            if (zone != null)
            {
                _context.Zones.Remove(zone);
                await _context.SaveChangesAsync();
            }
            return RedirectToAction(nameof(Index));
        }
        public IActionResult ExportToExcel()
        {
            var zones = _context.Zones.ToList();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Zones");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Name";
                worksheet.Cell(currentRow, 3).Value = "IsSelected";

                foreach (var zone in zones)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = zone.Id;
                    worksheet.Cell(currentRow, 2).Value = zone.Name;
                    worksheet.Cell(currentRow, 3).Value = zone.IsSelected;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "zones.xlsx");
                }
            }
        }
    }

    public class AdminViewModel
    {
        public IEnumerable<FormData> FormData { get; set; }
        public IEnumerable<Pracodawca> Pracodawcy { get; set; }
        public IEnumerable<Zone> Zones { get; set; }
    }
}
