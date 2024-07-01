using System;
using System.ComponentModel.DataAnnotations;

namespace web222.Models
{
    public class FormData
    {
        public int Id { get; set; }
        
        [Required]
        public string Nazwisko { get; set; }
        
        [Required]
        public string Imie { get; set; }
        
        public string Pracodawca { get; set; }
        
        public DateTime? DataWeryfikacji { get; set; }
        
        public DateTime? DataWyslaniaDoSG { get; set; }
        
        public string NumerWNP { get; set; }
        
        public string WNP { get; set; }
        
        public string KierownikPrzemek { get; set; }
        
        public DateTime? DataOdbioruID { get; set; }
        
        public string Uwagi { get; set; }
        
        public string DzialBB { get; set; }
    }
}