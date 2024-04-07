using Models;
using System.Data;

namespace Task3.Services
{
    public interface IExcelService
    {
        DataTable[] ImportExcelToDataTable();
        void UpdateClientsTable(DataTable clientsTable, Client client);
        void UpdateExcel(DataTable clientsTable);
    }
}
