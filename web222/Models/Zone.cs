using System.ComponentModel.DataAnnotations;

namespace web222.Models
{
    public class Zone
    {
        public int Id { get; set; }

        [Required]
        public string Name { get; set; }

        [Required]
        public bool IsSelected { get; set; }
    }
}