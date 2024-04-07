using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Models;
using System.Data;


namespace Task3.Services
{
    public class ExcelService : IExcelService
    {
        private readonly string _filePath;
        public ExcelService(string filePath) { 
            _filePath = filePath;
        }          

        public DataTable[] ImportExcelToDataTable()
        {
            DataTable[] table = new DataTable[3];
            
            using (XLWorkbook workBook = new XLWorkbook(_filePath))
            {               
                    IXLWorksheet workSheet;
                    for (int i = 1; i<=workBook.Worksheets.Count; i++)
                    {
                        workSheet = workBook.Worksheet(i);
                        switch (i)
                        {
                            default:
                                Console.WriteLine("Неверные данные");
                                break;
                            case 1:
                                DataTable dataTableProducts = ReadProductsWorksheet();
                                ReadTable(dataTableProducts, workSheet);
                                table[i-1] = dataTableProducts;
                                break;
                            case 2:
                                DataTable dataTableClients = ReadClientsWorksheet();
                                ReadTable(dataTableClients, workSheet);
                                table[i-1] = dataTableClients;
                                break;
                            case 3:
                                DataTable dataTableApplication = ReadApplicationsWorksheet();
                                ReadTable(dataTableApplication, workSheet);
                                table[i-1] = dataTableApplication;
                                break;
                        }
                    }               
            }
            return table;
        }

        public void UpdateClientsTable(DataTable clientsTable, Client client) {
            var rows = clientsTable.Rows;
            foreach (var item in rows) {
                DataRow row = item as DataRow;
                if (row[0].ToString() == client.Code.ToString())
                {
                    row[3] = client.ContactFace;
                }
            }        
        }

        public void UpdateExcel(DataTable clientsTable) {
            using (Stream stream = new FileStream(_filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)) {
                using (XLWorkbook workbook = new XLWorkbook(stream)) {
                    var workSheet = workbook.Worksheet(2);
                    if (workSheet != null) {
                        int rowNumber = 2;
                        int columnIndex = 1;
                        foreach (var row in clientsTable.Rows)
                        {
                            DataRow item = row as DataRow;
                            workSheet.Cell(rowNumber, columnIndex++).Value = item[0].ToString();
                            workSheet.Cell(rowNumber, columnIndex++).Value = item[1].ToString();
                            workSheet.Cell(rowNumber, columnIndex++).Value = item[2].ToString();
                            workSheet.Cell(rowNumber, columnIndex++).Value = item[3].ToString();
                            rowNumber++;
                            columnIndex=1;
                        }
                    }
                    workbook.Save();
                }
            }
        }

        private void ReadTable(DataTable dataTable, IXLWorksheet workSheet) {
            bool isFirstRow = true;
            foreach (var row in workSheet.Rows())
            {
                if (isFirstRow)
                {
                    isFirstRow = false;
                    continue;
                }
                else
                {
                    dataTable.Rows.Add();
                    int k = 0;
                    foreach (IXLCell cell in row.Cells())
                    {
                        if (k>dataTable.Columns.Count-1)
                        {
                            break;
                        }
                        dataTable.Rows[dataTable.Rows.Count - 1][k] = cell.Value.ToString();
                        k++;
                    }
                }
            }
        }
        private DataTable ReadClientsWorksheet() {
            DataTable dt = new DataTable();
            DataColumn dataColumn = new DataColumn("Код клиента");
            dt.Columns.Add(dataColumn);
            dataColumn=new DataColumn("Наименование организации");
            dt.Columns.Add(dataColumn);
            dataColumn=new DataColumn("Адрес");
            dt.Columns.Add(dataColumn);
            dataColumn=new DataColumn("Контактное лицо (ФИО)");
            dt.Columns.Add(dataColumn);
            return dt;
        }
        private DataTable ReadProductsWorksheet()
        {
            DataTable dt = new DataTable();
            DataColumn dataColumn = new DataColumn("Код товара");
            dt.Columns.Add(dataColumn);
            dataColumn=new DataColumn("Наименование");
            dt.Columns.Add(dataColumn);
            dataColumn=new DataColumn("Ед.измерения");
            dt.Columns.Add(dataColumn);
            dataColumn=new DataColumn("Цена товара за единицу");
            dt.Columns.Add(dataColumn);
            return dt;
        }
        private DataTable ReadApplicationsWorksheet()
        {
            DataTable dt = new DataTable();
            DataColumn dataColumn = new DataColumn("Код заявки");
            dt.Columns.Add(dataColumn);
            dataColumn=new DataColumn("Код товара");
            dt.Columns.Add(dataColumn);
            dataColumn=new DataColumn("Код клиента");
            dt.Columns.Add(dataColumn);
            dataColumn=new DataColumn("Номер заявки");
            dt.Columns.Add(dataColumn);
            dataColumn=new DataColumn("Требуемое количество");
            dt.Columns.Add(dataColumn);
            dataColumn=new DataColumn("Дата размещения");
            dt.Columns.Add(dataColumn);
            return dt;
        }


    }
}
