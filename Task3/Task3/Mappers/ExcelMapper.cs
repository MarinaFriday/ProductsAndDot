using Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Task3.Mappers
{
    internal class ExcelMapper : IExcelMapper
    {
        public Application MapToApplication(DataRow row)
        {
            int? code = null;
            if (Int32.TryParse(row.ItemArray[0].ToString(), out int resultCode)) code=(int?)resultCode;

            int? productCode = null;
            if (Int32.TryParse(row.ItemArray[1].ToString(), out int resultProductCode)) productCode=(int?)resultProductCode;

            int? clientCode = null;
            if (Int32.TryParse(row.ItemArray[2].ToString(), out int resultClientCode)) clientCode=(int?)resultClientCode;

            int? number = null;
            if (Int32.TryParse(row.ItemArray[3].ToString(), out int resultNumber)) number=(int?)resultNumber;

            int? count = null;
            if (Int32.TryParse(row.ItemArray[4].ToString(), out int resultCount)) count=(int?)resultCount;

            DateTime? date = null;
            if (DateTime.TryParse(row.ItemArray[5].ToString(), out DateTime resultDateTime)) date=(DateTime?)resultDateTime;

            return new Application() {
                Code = code,
                ProductCode = productCode,
                ClientCode = clientCode,
                Number = number,
                Count = count,
                Date = date,
            };
        }

        public Client MapToClient(DataRow row)
        {
            int? code = null;
            if (Int32.TryParse(row.ItemArray[0].ToString(), out int result)) code=(int?)result;
            
            return new Client()
            {
                Code = code,
                OrganisationName = (string?)row.ItemArray[1],
                Address = (string?)row.ItemArray[2],
                ContactFace = (string?)row.ItemArray[3]
            };
        }

        public Product MapToProduct(DataRow row)
        {
            int? code = null;
            if (Int32.TryParse(row.ItemArray[0].ToString(), out int resultInt)) code=(int?)resultInt;

            decimal? price = null;
            if (Decimal.TryParse(row.ItemArray[3].ToString(), out decimal resultDecimal)) price=(decimal?)resultDecimal;

            return new Product()
            {
                Code = code,
                Name = (string?)row.ItemArray[1],
                Unit = (string?)row.ItemArray[2],
                Price = price
            };
        }        
       

    }
}
