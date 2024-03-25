using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace PracticalTask3
{
    public class Product
    {
        public int ProductID { get; set; }
        public string ProductName { get; set; }
        public string UnitOfMeasurement { get; set; }
        public decimal Price { get; set; }
    }

    public class Client
    {
        public int ClientID { get; set; }
        public string OrganizationName { get; set; }
        public string Address { get; set; }
        public string ContactPerson { get; set; }
    }

    public class Order
    {
        public int OrderID { get; set; }
        public int ProductID { get; set; }
        public int ClientID { get; set; }
        public int OrderNumber { get; set; }
        public int Quantity { get; set; }
        public DateTime Date { get; set; }
    }
    internal class Program
    {
        public static List<Product> products;
        public static List<Client> clients;
        public static List<Order> orders;
        public static string? filePath;

        public static void LoadDataFromExcel()
        {
            using (XLWorkbook workbook = new XLWorkbook(filePath))
            {
                IXLWorksheet productsWorksheet = workbook.Worksheet("Товары");
                IXLWorksheet clientsWorksheet = workbook.Worksheet("Клиенты");
                IXLWorksheet ordersWorksheet = workbook.Worksheet("Заявки");

                foreach (var row in productsWorksheet.RowsUsed().Skip(1))
                {
                    Product product = new Product();
                    product.ProductID = int.Parse(row.Cell(1).Value.ToString());
                    product.ProductName = row.Cell(2).Value.ToString();
                    product.UnitOfMeasurement = row.Cell(3).Value.ToString();
                    product.Price = decimal.Parse(row.Cell(4).Value.ToString());

                    products.Add(product);
                }

                foreach (var row in clientsWorksheet.RowsUsed().Skip(1))
                {
                    Client client = new Client();
                    client.ClientID = int.Parse(row.Cell(1).Value.ToString());
                    client.OrganizationName = row.Cell(2).Value.ToString();
                    client.Address = row.Cell(3).Value.ToString();
                    client.ContactPerson = row.Cell(4).Value.ToString();

                    clients.Add(client);
                }

                foreach (var row in ordersWorksheet.RowsUsed().Skip(1))
                {
                    Order order = new Order();
                    order.OrderID = int.Parse(row.Cell(1).Value.ToString());
                    order.ProductID = int.Parse(row.Cell(2).Value.ToString());
                    order.ClientID = int.Parse(row.Cell(3).Value.ToString());
                    order.OrderNumber = int.Parse(row.Cell(4).Value.ToString());
                    order.Quantity = int.Parse(row.Cell(5).Value.ToString());
                    order.Date = DateTime.Parse(row.Cell(6).Value.ToString());

                    orders.Add(order);
                }
            }
        }

        public static List<Order> GetOrdersByProductName(string productName)
        {
            List<Order> result = new List<Order>();

            foreach (Order order in orders)
            {
                int productID = products.Find(p => p.ProductName == productName).ProductID;

                if (order.ProductID == productID)
                {
                    result.Add(order);
                }
            }

            return result;
        }

        public static bool ChangeClientContactPerson(string organizationName, string contactPerson)
        {
            Client client = clients.Find(c => c.OrganizationName == organizationName);

            if (client != null)
            {
                client.ContactPerson = contactPerson;

                using (XLWorkbook workbook = new XLWorkbook(filePath))
                {
                    IXLWorksheet clientsWorksheet = workbook.Worksheet("Клиенты");
                    clientsWorksheet.Cell("A1").Value = "Код клиента";
                    clientsWorksheet.Cell("B1").Value = "Наименование организации";
                    clientsWorksheet.Cell("C1").Value = "Адрес";
                    clientsWorksheet.Cell("D1").Value = "Контактное лицо (ФИО)";
                    int i = 2;
                    foreach(Client writeClient in clients)
                    {
                        clientsWorksheet.Cell("A"+i).Value = writeClient.ClientID;
                        clientsWorksheet.Cell("B"+i).Value = writeClient.OrganizationName;
                        clientsWorksheet.Cell("C"+i).Value = writeClient.Address;
                        clientsWorksheet.Cell("D"+i).Value = writeClient.ContactPerson;
                        i++;
                    }
                    workbook.Save();
                }

                return true;
            }

            return false;
        }

        public static Client GetClientWithMostOrders(int year, int month)
        {
            Dictionary<int, int> ordersCountByClient = new Dictionary<int, int>();

            foreach (Order order in orders)
            {
                if (order.Date.Year == year && order.Date.Month == month)
                {
                    if (ordersCountByClient.ContainsKey(order.ClientID))
                    {
                        ordersCountByClient[order.ClientID]++;
                    }
                    else
                    {
                        ordersCountByClient.Add(order.ClientID, 1);
                    }
                }
            }

            int clientIDWithMostOrders = 0;
            int maxOrdersCount = 0;

            foreach (var item in ordersCountByClient)
            {
                if (item.Value > maxOrdersCount)
                {
                    clientIDWithMostOrders = item.Key;
                    maxOrdersCount = item.Value;
                }
            }

            Client clientWithMostOrders = clients.Find(c => c.ClientID == clientIDWithMostOrders);

            return clientWithMostOrders;
        }
        static void Main(string[] args)
        {
            products = new List<Product>();
            clients = new List<Client>();
            orders = new List<Order>();
            filePath = "";
            string exercise = "";
            bool work = true;
            while (work)
            {
                try
                {
                    Console.WriteLine("==============================");
                    Console.WriteLine("Меню:\n1) Ввести путь к файлу с данными\n2) Найти заказы по наименованию товара\n" +
                        "3) Изменить контактное лицо у организации\n4) Определить золотого клиента\n0) Выход из программы");
                    exercise = Console.ReadLine();
                    switch (exercise)
                    {
                        case "1":
                            try
                            {
                                Console.WriteLine("Введите путь до Excel файла:");
                                filePath = "C:\\Users\\Alex\\Downloads\\Практическое задание для кандидата.xlsx";
                                //filePath = Console.ReadLine();
                                LoadDataFromExcel();
                            }
                            catch
                            {
                                   Console.WriteLine("Некорректное название или структура файла.");
                            }
                            break;

                        case "2":
                            try
                            {
                                Console.WriteLine("Введите наименование товара:");
                                string productName = Console.ReadLine();
                                List<Order> ordersByProductName = GetOrdersByProductName(productName);
                                foreach (Order order in ordersByProductName)
                                {
                                    Console.WriteLine("-----------");
                                    Console.WriteLine($"Клиент: {clients.Find(c => c.ClientID == order.ClientID).OrganizationName}");
                                    Console.WriteLine($"Количество товара: {order.Quantity}");
                                    Console.WriteLine($"Цена товара: {products.Find(p => p.ProductID == order.ProductID).Price}");
                                    Console.WriteLine($"Дата заказа: {order.Date.ToShortDateString()}");
                                    Console.WriteLine("-----------");
                                }
                            }
                            catch 
                            {
                                Console.WriteLine("Товар с указанным наименованием отсутствует в базе.");
                            }
                            break;

                        case "3":
                            try
                            {
                                Console.WriteLine("Введите название организации клиента:");
                                string organizationName = Console.ReadLine();
                                Console.WriteLine("Введите ФИО нового контактного лица:");
                                string contactPerson = Console.ReadLine();

                                bool result = ChangeClientContactPerson(organizationName, contactPerson);

                                if (result)
                                {
                                    Console.WriteLine("Изменения сохранены успешно.");
                                }
                                else
                                {
                                    Console.WriteLine("Не удалось найти клиента с указанной организацией.");
                                }
                            }
                            catch
                            {
                                Console.WriteLine("Указанная организация отсутствует в базе данных.");
                            }
                            break;

                        case "4":
                            try
                            {
                                Console.WriteLine("Введите год:");
                                int year = int.Parse(Console.ReadLine());
                                Console.WriteLine("Введите месяц:");
                                int month = int.Parse(Console.ReadLine());

                                Client clientWithMostOrders = GetClientWithMostOrders(year, month);

                                Console.WriteLine($"Клиент с наибольшим количеством заказов: {clientWithMostOrders.OrganizationName}");
                            }
                            catch (FormatException)
                            {
                                Console.WriteLine("Вы ввели некорректные данные");
                            }
                            catch 
                            {
                                Console.WriteLine("В указанные месяц и год нет заказов.");
                            }
                            break;

                        case "0":
                            work = false;
                            break;

                        default:
                            Console.WriteLine("Неизвестная команда");
                            break;
                    }
                }
                catch
                {
                    Console.WriteLine("Что-то пошло не так. Вероятно имеется ошибка в данных из файла или вводимых данных.");
                }
            }
        }
    }
}
