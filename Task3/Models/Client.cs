using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class Client
    {
        public int? Code { get; set; }
        public string? OrganisationName { get; set; }
        public string? Address { get; set; }
        public string? ContactFace { get; set; }


        public override string ToString()
        {
            return $"{Code} {OrganisationName} {Address} {ContactFace}";
        }
    }
}
