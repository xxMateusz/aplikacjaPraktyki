using System.ComponentModel.DataAnnotations;

namespace web222.Models
{
    public class Pracodawca
    {
        public int Id { get; set; }

        [Required]
        public string Nazwa { get; set; }
    }
}