using System.Data;
using Models;

namespace Task3.Mappers
{
    internal interface IExcelMapper
    {
        Application MapToApplication(DataRow row);
        Client MapToClient(DataRow row);
        Product MapToProduct(DataRow row);
    }
}
