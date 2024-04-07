using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Spreadsheet;
using Models;
using System.Data;
using Task3.Mappers;
using Task3.Services;
using static System.Runtime.InteropServices.JavaScript.JSType;

Console.WriteLine("Введите путь до файла: ");
//"practice.xlsx"
var filePath = Console.ReadLine();

IExcelMapper mapper = new ExcelMapper();
IExcelService excelService = new ExcelService(filePath);


DataTable[] tables = new DataTable[3];
tables = excelService.ImportExcelToDataTable();

List<Client> clients = new List<Client>();
List<Product> products = new List<Product>();
List<Application> applications = new List<Application>();

bool menu = true;

do{
    Console.WriteLine("------Меню------");
    Console.WriteLine("1. Вывести информацию файла в консоль");
    Console.WriteLine("2. По наименованию товара выводить " +
                      "информацию о клиентах, заказавших этот товар, с \n" +
                      "указанием информации по количеству товара, цене и дате заказа");
    Console.WriteLine("3. Обновление контактного лица");
    Console.WriteLine("4. Запрос на определение золотого клиента, клиента с наибольшим количеством заказов," +
                      "\nза указанный год, месяц");
    Console.WriteLine("5. Выход");
    Console.WriteLine();

    int choice = int.Parse(Console.ReadLine());
    MappingAndConsole();

    switch (choice)
    {
        default:
            Console.WriteLine("Неверный ввод");
            break;
        case 1:
            Console.Clear();
            MappingAndConsole();
            break;
        case 2:
            Console.Clear();
            Console.WriteLine("Введине наименование товара");
            string nameProduct = Console.ReadLine().Trim();
            InformationByProductName(nameProduct);
            break;
        case 3:
            Console.Clear();
            Console.WriteLine("Введите название организации: ");
            string nameOrganisation = Console.ReadLine().Trim();
            Console.WriteLine("Введите ФИО контактного лица (через пробелы): ");
            string contactFace = Console.ReadLine().Trim();
            UpdateClient(nameOrganisation, contactFace);
            excelService.UpdateExcel(tables[1]);
            Console.WriteLine("Данные изменены");
            break;
        case 4:
            Console.Clear();       
            Console.WriteLine("1. За год");
            Console.WriteLine("2. За месяц");
            int c = int.Parse(Console.ReadLine());
            switch (c) {
                default:
                    Console.WriteLine("Неверный ввод");
                    break;
                case 1:
                    Console.WriteLine("Укажите год");
                    int year = int.Parse(Console.ReadLine());
                    GoldClient(true, year);
                    break;
                case 2:
                    Console.WriteLine("Укажите месяц");
                    int month = int.Parse(Console.ReadLine());
                    GoldClient(false, month);
                    break;
            }
            break;
        case 5:
            menu=false;
            break;
    }
} while (menu);

void InformationByProductName(string nameProduct) {
    
    var selectApplication = from a in applications
                            join p in products on a.ProductCode equals p.Code
                            join c in clients on a.ClientCode equals c.Code
                            where nameProduct == p.Name
                            select new
                            {
                                NameProduct = nameProduct,
                                ClientCode = c.Code,
                                OrganisationName = c.OrganisationName,
                                Address = c.Address,
                                ContactFace = c.ContactFace,
                                CountProducts = a.Count,
                                Price = p.Price,
                                DateOrder = a.Date
                            };
    InformationOnConsole(selectApplication.ToList());
}

void UpdateClient(string nameOrganisation, string contactFace) {
    var selectClient = (from c in clients
                       where nameOrganisation == c.OrganisationName
                       select new
                       {
                           Code = c.Code,
                           OrganisationName = c.OrganisationName,
                           Address = c.Address,
                           ContactFace = contactFace
                       }).FirstOrDefault();
    
    var clientEdit = clients.FirstOrDefault(c => c.Code == selectClient.Code);
    if (clientEdit != null)
    {
        clientEdit.ContactFace = selectClient.ContactFace;
        InformationOnConsole<Client>(clients);
        excelService.UpdateClientsTable(tables[1], clientEdit);
    }
}

void MappingAndConsole()
{    
    for (int i = 0; i<tables.Length; i++)
    {
        switch (i)
        {
            default:
                Console.WriteLine($"{i}");
                break;
            case 0:
                foreach (DataRow dataRow in tables[i].Rows)
                {
                    if (dataRow.ItemArray.Any(e => e.ToString()=="")) continue;
                    products.Add(mapper.MapToProduct(dataRow));
                }
                InformationOnConsole<Product>(products);
                break;
            case 1:
                foreach (DataRow dataRow in tables[i].Rows)
                {
                    if (dataRow.ItemArray.Any(e => e.ToString()=="")) continue;
                    clients.Add(mapper.MapToClient(dataRow));
                }
                InformationOnConsole<Client>(clients);
                break;
            case 2:
                foreach (DataRow dataRow in tables[i].Rows)
                {
                    if (dataRow.ItemArray.Any(e => e.ToString()=="")) continue;
                    applications.Add(mapper.MapToApplication(dataRow));
                }
                InformationOnConsole<Application>(applications);
                break;
        }
    }
}

void InformationOnConsole<T>(List<T> list) {
    foreach (var item in list)
    {
        Console.WriteLine(item?.ToString());
    }
}

void GoldClient(bool isYear, int number) {
    if (isYear) {
        var selectClient = (from c in clients
                            join a in applications on c.Code equals a.ClientCode
                            where a.Date.Value.Year == number
                            group c by c.Code into g
                            orderby g.Count() descending
                            select new
                            {
                                Code = (from c in g select c.Code).FirstOrDefault(),
                                OrganisationName = (from c in g select c.OrganisationName).FirstOrDefault(),
                                Address = (from c in g select c.Address).FirstOrDefault(),
                                ContactFace = (from c in g select c.ContactFace).FirstOrDefault()
                            }).FirstOrDefault();
        Console.WriteLine(selectClient);
    } else
    {
        var selectClient = (from c in clients
                            join a in applications on c.Code equals a.ClientCode
                            where a.Date.Value.Month == number
                            group c by c.Code into g
                            orderby g.Count() descending
                            select new
                            {
                                Code = (from c in g select c.Code).FirstOrDefault(),
                                OrganisationName = (from c in g select c.OrganisationName).FirstOrDefault(),
                                Address = (from c in g select c.Address).FirstOrDefault(),
                                ContactFace = (from c in g select c.ContactFace).FirstOrDefault()
                            }).FirstOrDefault();
        Console.WriteLine(selectClient);
    }

}




