using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class Application
    {
        public int? Code { get; set; }
        public int? ProductCode { get; set; }
        public int? ClientCode { get; set; }
        public int? Number { get; set; }
        public int? Count { get; set; }
        public DateTime? Date { get; set; }

        public override string ToString()
        {
            return $"{Code} {ProductCode} {ClientCode} {Number} {Count} {Date.Value.ToShortDateString()}";
        }
    }
}
