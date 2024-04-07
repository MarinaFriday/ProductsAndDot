using System.Net;

namespace Models
{
    public class Product
    {
        public int? Code { get; set; }
        public string? Name { get; set; }
        public string? Unit { get; set; }
        public decimal? Price { get; set; }

        public override string ToString()
        {
            return $"{Code} {Name} {Unit} {Price}";
        }
    }
}
