using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using OfficeOpenXml;
using System.IO;
using ClosedXML.Excel;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using System;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word; // Добавьте эту директиву
using System.IO;
using System.Diagnostics;

namespace OrderWorkAuto
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        private decimal serviceCost = 0;
        private decimal detailsCost = 0;
        private decimal originalServiceCost = 0;
        private decimal originalDetailsCost = 0;
        private bool isCargoCar = false; // Переменная для отслеживания типа автомобиля
        private string templatePath = "ИПЗаказ.docx";
        private string copyPathField = "Копия_ИПЗаказ.docx"; // Переименовали поле класса, чтобы не скрывать локальную переменную copyPath

        public MainWindow()
        {
            InitializeComponent();
            DateTime? selectedTime = DateTime.Now;
            TimePicker.Value = selectedTime;
            selectedDetails.CollectionChanged += SelectedDetails_CollectionChanged;
            DiscountTextBox.Text = "0";
            CarTypeRadioButton1.IsChecked = true;
            LoadDataFromExcel();
            LoadDataFromExcelDetails();
            LoadCarMarksFromExcel();
            LoadCarModelsFromExcel();
            LoadDataFromExcelCounterAgents();
            DiscountDetailBox.Text = "0";
            DiscountServBox.Text = "0";
            ReadValueFromFileAndUpdateTextBox();
            // Получение пути к файлу
            

        }
        private ObservableCollection<Service> selectedServices = new ObservableCollection<Service>();
        private ObservableCollection<Detail> selectedDetails = new ObservableCollection<Detail>();
        private List<string> carMarks = new List<string>();
      

        public class Service : INotifyPropertyChanged
        {
            private string serviceName;
            public string ServiceName
            {
                get { return serviceName; }
                set
                {
                    if (serviceName != value)
                    {
                        serviceName = value;
                        OnPropertyChanged(nameof(ServiceName));
                    }
                }
            }
            private decimal price;
            public decimal Price
            {
                get { return price; }
                set
                {
                    if (price != value)
                    {
                        price = value;
                        OnPropertyChanged(nameof(Price));
                        OnPropertyChanged(nameof(TotalCost));
                    }
                }
            }

            private decimal cost;
            public decimal Cost
            {
                get { return cost; }
                set
                {
                    if (cost != value)
                    {
                        cost = value;
                        OnPropertyChanged(nameof(Cost));
                        OnPropertyChanged(nameof(TotalCost));
                    }
                }
            }

            private int count;
            public int Count
            {
                get { return count; }
                set
                {
                    if (count != value)
                    {
                        count = value;
                        OnPropertyChanged(nameof(Count));
                        OnPropertyChanged(nameof(TotalCost));
                    }
                }
            }
            private decimal gruzCost;
            public decimal GruzCost
            {
                get { return gruzCost; }
                set
                {
                    if (gruzCost != value)
                    {
                        gruzCost = value;
                        OnPropertyChanged(nameof(GruzCost));
                    }
                }
            }

            public decimal TotalCost => Cost * Count;
            public decimal OriginalCost { get; set; }

            public event PropertyChangedEventHandler PropertyChanged;

            protected void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        public class CounterAgent
        {
            public string AgentName { get; set; }
        }
        public class Detail : INotifyPropertyChanged
        {
            private string detalName;
            public string DetalName
            {
                get { return detalName; }
                set
                {
                    detalName = value;
                    OnPropertyChanged("DetalName");
                }
            }

            private decimal cost;
            public decimal Cost
            {
                get { return cost; }
                set
                {
                    cost = value;
                    OnPropertyChanged("Cost");
                    OnPropertyChanged("TotalCost");
                }
            }

            private int quantity;
            public int Quantity
            {
                get { return quantity; }
                set
                {
                    quantity = value;
                    OnPropertyChanged("Quantity");
                    OnPropertyChanged("TotalCost");
                }
            }

            public decimal TotalCost
            {
                get { return Cost * Quantity; }
            }

            private decimal gruzCost; // Добавленное поле для GruzCost
            public decimal GruzCost
            {
                get { return gruzCost; }
                set
                {
                    gruzCost = value;
                    OnPropertyChanged("GruzCost");
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;

            protected void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }

            public decimal OriginalCost { get; set; } // Поле для исходной цены
        }
        private void UpdateZakazNar()
        {
            string filePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "order_number.txt");

            try
            {
                if (File.Exists(filePath))
                {
                    // Чтение содержимого файла
                    string[] lines = File.ReadAllLines(filePath);
                    int currentOrderNumber = 0;

                    // Изменение значения OrderNumber
                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (lines[i].StartsWith("OrderNumber="))
                        {
                            string[] parts = lines[i].Split('=');
                            if (parts.Length == 2 && int.TryParse(parts[1], out currentOrderNumber))
                            {
                                currentOrderNumber++; // Увеличиваем значение на единицу
                                lines[i] = $"OrderNumber={currentOrderNumber}"; // Обновляем строку с новым значением
                                break;
                            }
                        }
                    }

                    // Запись измененных данных обратно в файл
                    File.WriteAllLines(filePath, lines);

                    // Обновление значения OrderNumberBox
                    OrderNumberBox.Text = currentOrderNumber.ToString();
                }
                else
                {
                    // Если файл не найден, можно вывести сообщение об ошибке или создать новый файл с начальным значением OrderNumber=1
                    File.WriteAllText(filePath, "OrderNumber=1");
                    OrderNumberBox.Text = "1";
                }
            }
            catch (Exception ex)
            {
                // Обработка ошибок при чтении файла
                OrderNumberBox.Text = "Ошибка: " + ex.Message;
            }
        }
        private void ReadValueFromFileAndUpdateTextBox()
        {
            string filePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "order_number.txt");
            try
            {
                if (File.Exists(filePath))
                {
                    // Чтение содержимого файла
                    string[] lines = File.ReadAllLines(filePath);
                    string value = "";

                    // Поиск строки с нужным значением
                    foreach (string line in lines)
                    {
                        if (line.StartsWith("OrderNumber=")) // Измените "Value=" на нужную вам метку или формат строки
                        {
                            // Извлечение значения из строки
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                value = parts[1];
                                break;
                            }
                        }
                    }

                    // Установка считанного значения в TextBox
                    OrderNumberBox.Text = value;
                }
                else
                {
                    // Если файл не найден, можно вывести сообщение об ошибке или установить другое значение по умолчанию в TextBox
                    OrderNumberBox.Text = "Файл не найден";
                }
            }
            catch (Exception ex)
            {
                // Обработка ошибок при чтении файла
                OrderNumberBox.Text = "Ошибка: " + ex.Message;
            }
        }
        private void LoadCarMarksFromExcel()
        {
            // Установите путь к файлу Excel
            string excelFilePath = "ServicesList.xlsx"; // Путь к файлу относительно текущей директории

            try
            {
                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    var worksheet = workbook.Worksheet("List3"); // Имя вашего листа

                    foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропустить первую строку с заголовками
                    {
                        var carMark = row.Cell(1).Value.ToString();
                        carMarks.Add(carMark);
                    }

                    // Заполните ComboBox данными из списка марок автомобилей
                    MarksAutos.ItemsSource = carMarks;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}");
            }
        }
        private void LoadCarModelsFromExcel()
        {
            // Замените это на соответствующий путь к файлу Excel
            string excelFilePath = "ServicesList.xlsx";

            try
            {
                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    var worksheet = workbook.Worksheet("List4"); // Имя вашего листа

                    List<string> models = new List<string>();

                    foreach (var cell in worksheet.Column(1).CellsUsed().Skip(1)) // Пропустить первую строку с заголовками
                    {
                        string model = cell.GetString(); // Получите значение ячейки

                        if (!string.IsNullOrEmpty(model))
                        {
                            models.Add(model);
                        }
                    }

                    // Заполните ComboBox данными из списка models
                    ModelAutos.ItemsSource = models;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}");
            }
        }

        private void LoadDataFromExcel()
        {
            string excelFilePath = "ServicesList.xlsx"; // Путь к файлу относительно текущей директории
            if(CarTypeRadioButton1.IsChecked == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook(excelFilePath))
                    {
                        var worksheet = workbook.Worksheet("List1"); // Имя вашего листа

                        System.Data.DataTable dt = new System.Data.DataTable();
                        dt.Columns.Add("ServiceName");
                        dt.Columns.Add("Cost", typeof(string));
                        dt.Columns.Add("Count", typeof(string)); // Добавьте столбец "Count" как string
                        dt.Columns.Add("GruzCost", typeof(string));
                        foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропустить первую строку с заголовками
                        {
                            var serviceName = row.Cell(1).Value.ToString();
                            var cost = row.Cell(3).Value.ToString(); // Сохраните стоимость как строку


                            if (decimal.TryParse(cost, out decimal parsedCost))

                            {
                                dt.Rows.Add(serviceName, cost);

                                // Найдите строку "Цена" и сохраните её стоимость и gruzCost
                                if (serviceName == "Цена")
                                {
                                    originalServiceCost = parsedCost;
                                    // Предположим, что "Цена" в Excel соответствует услуге "Цена"
                                    Service priceService = new Service
                                    {
                                        ServiceName = "Цена",
                                        Cost = parsedCost,

                                    };
                                    // Добавляем созданный элемент в ComboBox
                                    dt.Rows.Add(priceService.ServiceName, priceService.Cost.ToString(), priceService.GruzCost.ToString());
                                }
                            }
                        }

                        // Заполните ComboBox данными из DataTable
                        Services.ItemsSource = dt.DefaultView;

                        // Вставьте этот код для настройки шаблона элементов ComboBox
                        Services.DisplayMemberPath = null; // Очистите настройку отображаемого пути
                        Services.SelectedValuePath = null; // Очистите настройку значения

                        // Создайте шаблон для элементов ComboBox
                        Services.ItemTemplate = new DataTemplate();

                        var textBlock = new FrameworkElementFactory(typeof(TextBlock));
                        textBlock.SetBinding(TextBlock.TextProperty, new Binding("ServiceName"));
                        textBlock.SetBinding(TextBlock.TagProperty, new Binding("Cost"));


                        Services.ItemTemplate.VisualTree = textBlock;
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}");
                }
            }
            if(CarTypeRadioButton2.IsChecked == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook(excelFilePath))
                    {
                        var worksheet = workbook.Worksheet("List1"); // Имя вашего листа

                        System.Data.DataTable dt = new System.Data.DataTable();
                        dt.Columns.Add("ServiceName");
                        dt.Columns.Add("Cost", typeof(string));
                        dt.Columns.Add("Count", typeof(string)); // Добавьте столбец "Count" как string
                        dt.Columns.Add("GruzCost", typeof(string));
                        foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропустить первую строку с заголовками
                        {
                            var serviceName = row.Cell(1).Value.ToString();
                            var cost = row.Cell(4).Value.ToString(); // Сохраните стоимость как строку
                           

                            if (decimal.TryParse(cost, out decimal parsedCost) )
                                
                            {
                                dt.Rows.Add(serviceName, cost);

                                // Найдите строку "Цена" и сохраните её стоимость и gruzCost
                                if (serviceName == "Цена")
                                {
                                    originalServiceCost = parsedCost;
                                    // Предположим, что "Цена" в Excel соответствует услуге "Цена"
                                    Service priceService = new Service
                                    {
                                        ServiceName = "Цена",
                                        Cost = parsedCost,
                                       
                                    };
                                    // Добавляем созданный элемент в ComboBox
                                    dt.Rows.Add(priceService.ServiceName, priceService.Cost.ToString(), priceService.GruzCost.ToString());
                                }
                            }
                        }

                        // Заполните ComboBox данными из DataTable
                        Services.ItemsSource = dt.DefaultView;

                        // Вставьте этот код для настройки шаблона элементов ComboBox
                        Services.DisplayMemberPath = null; // Очистите настройку отображаемого пути
                        Services.SelectedValuePath = null; // Очистите настройку значения

                        // Создайте шаблон для элементов ComboBox
                        Services.ItemTemplate = new DataTemplate();

                        var textBlock = new FrameworkElementFactory(typeof(TextBlock));
                        textBlock.SetBinding(TextBlock.TextProperty, new Binding("ServiceName"));
                        textBlock.SetBinding(TextBlock.TagProperty, new Binding("Cost"));


                        Services.ItemTemplate.VisualTree = textBlock;
                    }
                }
                
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}");
                }
            }
                
            
            
         

            
        }
        private void LoadDataFromExcelDetails()
        {
            string excelFilePath = "ServicesList.xlsx"; // Путь к файлу относительно текущей директории

            if (CarTypeRadioButton1.IsChecked == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook(excelFilePath))
                    {
                        var worksheet = workbook.Worksheet("List2"); // Имя вашего листа

                        System.Data.DataTable dt = new System.Data.DataTable();
                        dt.Columns.Add("DetalName");
                        dt.Columns.Add("Cost", typeof(string)); // Установите тип данных столбца как string
                        dt.Columns.Add("GruzCost", typeof(string)); // Добавьте столбец "GruzCost" как string

                        foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропустить первую строку с заголовками
                        {
                            var detalName = row.Cell(1).Value.ToString();
                            var cost = row.Cell(3).Value.ToString(); // Сохраните стоимость как строку
                           

                            if (decimal.TryParse(cost, out decimal parsedCost))
                            {
                                dt.Rows.Add(detalName, cost);

                                // Найдите строку "Цена" и сохраните её стоимость
                                if (detalName == "Цена")
                                {
                                    originalDetailsCost = parsedCost;
                                }
                            }
                        }

                        // Заполните ComboBox данными из DataTable
                        DetailsBox.ItemsSource = dt.DefaultView;

                        // Вставьте этот код для настройки шаблона элементов ComboBox
                        DetailsBox.DisplayMemberPath = null; // Очистите настройку отображаемого пути
                        DetailsBox.SelectedValuePath = null; // Очистите настройку значения

                        // Создайте шаблон для элементов ComboBox
                        DetailsBox.ItemTemplate = new DataTemplate();

                        var textBlock = new FrameworkElementFactory(typeof(TextBlock));
                        textBlock.SetBinding(TextBlock.TextProperty, new Binding("DetalName"));
                        textBlock.SetBinding(TextBlock.TagProperty, new Binding("Cost"));
                        textBlock.SetBinding(TextBlock.TagProperty, new Binding("GruzCost")); // Добавьте привязку для GruzCost

                        DetailsBox.ItemTemplate.VisualTree = textBlock;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}");
                }
            }
            if (CarTypeRadioButton2.IsChecked == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook(excelFilePath))
                    {
                        var worksheet = workbook.Worksheet("List2"); // Имя вашего листа

                        System.Data.DataTable dt = new System.Data.DataTable();
                        dt.Columns.Add("DetalName");
                        dt.Columns.Add("Cost", typeof(string)); // Установите тип данных столбца как string
                        dt.Columns.Add("GruzCost", typeof(string)); // Добавьте столбец "GruzCost" как string

                        foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропустить первую строку с заголовками
                        {
                            var detalName = row.Cell(1).Value.ToString();
                            var cost = row.Cell(4).Value.ToString(); // Сохраните стоимость как строку
                           

                            if (decimal.TryParse(cost, out decimal parsedCost))
                            {
                                dt.Rows.Add(detalName, cost);

                                // Найдите строку "Цена" и сохраните её стоимость
                                if (detalName == "Цена")
                                {
                                    originalDetailsCost = parsedCost;
                                }
                            }
                        }

                        // Заполните ComboBox данными из DataTable
                        DetailsBox.ItemsSource = dt.DefaultView;

                        // Вставьте этот код для настройки шаблона элементов ComboBox
                        DetailsBox.DisplayMemberPath = null; // Очистите настройку отображаемого пути
                        DetailsBox.SelectedValuePath = null; // Очистите настройку значения

                        // Создайте шаблон для элементов ComboBox
                        DetailsBox.ItemTemplate = new DataTemplate();

                        var textBlock = new FrameworkElementFactory(typeof(TextBlock));
                        textBlock.SetBinding(TextBlock.TextProperty, new Binding("DetalName"));
                        textBlock.SetBinding(TextBlock.TagProperty, new Binding("Cost"));
                        textBlock.SetBinding(TextBlock.TagProperty, new Binding("GruzCost")); // Добавьте привязку для GruzCost

                        DetailsBox.ItemTemplate.VisualTree = textBlock;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}");
                }
            }




        }
        private void LoadDataFromExcelCounterAgents()
        {
            // Установите путь к файлу Excel
            string excelFilePath = "ServicesList.xlsx"; // Путь к файлу относительно текущей директории

            try
            {
                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    var worksheet = workbook.Worksheet("CounterAgents"); // Имя листа с контрагентами

                    List<CounterAgent> counterAgents = new List<CounterAgent>();

                    foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропустить первую строку с заголовками
                    {
                        var agentName = row.Cell(1).Value.ToString(); // Первый столбец (A) содержит "AgentName"
                        counterAgents.Add(new CounterAgent { AgentName = agentName });
                    }

                    // Заполнить ComboBox данными из списка контрагентов
                    CounterAgentComboBox.ItemsSource = counterAgents;
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}");
            }
        }


        private void CreateOrder_Click(object sender, RoutedEventArgs e)
        {

        }
        private void UpdateTotalCostLabel()
        {
            // Получите значение скидки из обоих полей
            decimal serviceDiscount = 0;
            if (decimal.TryParse(DiscountServBox.Text, out serviceDiscount))
            {
                serviceDiscount /= 100; // Переведите проценты в десятичное значение
            }

            decimal detailDiscount = 0;
            if (decimal.TryParse(DiscountDetailBox.Text, out detailDiscount))
            {
                detailDiscount /= 100; // Переведите проценты в десятичное значение
            }

            // Обновите стоимость услуг с учетом скидки
            decimal serviceCost = selectedServices.Sum(service => service.Cost * service.Count) * (1 - serviceDiscount);
            string formattedServiceCost = serviceCost.ToString("0.00");

            // Обновите стоимость деталей с учетом скидки
            decimal detailsCost = selectedDetails.Sum(detail => detail.TotalCost) * (1 - detailDiscount);
            string formattedDetailsCost = detailsCost.ToString("0.00");

            // Обновите SummLabel и SummLabel2
            SummLabel.Content = $" {formattedServiceCost} руб.";
            SummLabel2.Content = $"{formattedDetailsCost} руб.";

            // Обновите общую стоимость с учетом скидок
            decimal totalCost = serviceCost + detailsCost;
            decimal discountedTotalCost = totalCost;

            TotalCostLabel.Content = $"{discountedTotalCost:N2} ₽";
        }
        private void Button_Click_AddService(object sender, RoutedEventArgs e)
        {
            if (Services.SelectedItem != null)
            {
                var selectedService = (DataRowView)Services.SelectedItem;
                var serviceName = selectedService["ServiceName"].ToString();
                var costString = selectedService["Cost"].ToString();

                if (decimal.TryParse(costString, out decimal cost))
                {
                    var service = new Service
                    {
                        ServiceName = serviceName,
                        Cost = cost,
                        Count = 1 // Установите количество в 1
                    };

                    // Проверка наличия услуги в списке
                    if (!selectedServices.Any(s => s.ServiceName == service.ServiceName))
                    {
                        service.OriginalCost = service.Cost; // Сохранить исходную цену
                        selectedServices.Add(service);

                        // Обновление DataGrid для отображения выбранных услуг
                        ServicesGrid.ItemsSource = null;
                        ServicesGrid.ItemsSource = selectedServices;

                        // Обновление Label с общей стоимостью
                        UpdateTotalCostServ();
                       // RecalculatePrices();
                        UpdateTotalCostLabel();
                    }
                }
            }
        }
        private void ServicesGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            var cell = e.Column.GetCellContent(e.Row);
            if (cell is TextBlock)
            {
                ((TextBlock)cell).Text = ""; // Удаление значения при начале редактирования
            }
        }
        private void ServicesGrid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            if (e.Column.Header.ToString() == "Цена") // Замените "Цена" на заголовок вашего столбца
            {
                TextBox textBox = e.EditingElement as TextBox;
                if (textBox != null)
                {
                    textBox.Text = ""; // Очистка содержимого ячейки при начале редактирования
                }
            }
        }
        private void ServicesGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                if (e.Column.Header.ToString() == "Количество")
                {
                    var editedService = (Service)e.Row.Item;

                    if (int.TryParse(((TextBox)e.EditingElement).Text, out int newCount))
                    {
                        editedService.Count = newCount;
                        UpdateTotalCostLabel();
                    }
                }
               
            }
            if (e.EditAction == DataGridEditAction.Commit)
            {
                if (e.Column.Header.ToString() == "Цена")
                {
                    var editedService = (Service)e.Row.Item;
                    var editedColumn = e.Column;

                    TextBox textBox = e.EditingElement as TextBox;
                    if (textBox != null && editedService != null && editedColumn != null && editedColumn.Header.ToString() == "Цена")
                    {
                        if (!string.IsNullOrWhiteSpace(textBox.Text) && decimal.TryParse(textBox.Text, out decimal newCost))
                        {
                            editedService.Cost = newCost;
                            UpdateTotalCostLabel();
                        }
                    }
                }
            }

        }



        private void DiscountTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(DiscountTextBox.Text))
            {
                UpdateTotalCostLabel();
                UpdateTotalCostLabel2();
            }
            
            
                UpdateTotalCostLabel();
                UpdateTotalCostLabel2();
            

        }
        private void DiscountServBox_TextChanged(object sender, TextChangedEventArgs e)
        {


           


            UpdateTotalCostServ();
            UpdateTotalCostLabel();



        }
        private void DiscountDetailBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            


            
            UpdateTotalCostDetail();
            UpdateTotalCostLabel();

        }
        private void UpdateTotalCostServ()
        {
            // Получите значение скидки из DiscountTextBox
            decimal discount = 0;
            if (decimal.TryParse(DiscountServBox.Text, out discount))
            {
                discount /= 100; // Переведите проценты в десятичное значение

                // Обновите TotalCostLabel
                decimal totalservcost = selectedServices.Sum(service => service.Cost * service.Count);
                decimal totalservcostwithdiscount = totalservcost - totalservcost * discount;
                SummLabel.Content = $"{totalservcostwithdiscount:N2} ₽";
                decimal totalCost =totalservcostwithdiscount + selectedDetails.Sum(detail => detail.TotalCost);
                decimal discountedTotalCost = totalCost;

                TotalCostLabel.Content = $"{discountedTotalCost:N2} ₽";
            }
        }
        private void UpdateTotalCostDetail()
        {
            // Получите значение скидки из DiscountTextBox
            decimal discount = 0;
            if (decimal.TryParse(DiscountDetailBox.Text, out discount))
            {
                discount /= 100; // Переведите проценты в десятичное значение
                decimal totaldetailcost = selectedDetails.Sum(detail => detail.TotalCost);
                decimal totaldetailcostwithdiscount = totaldetailcost - totaldetailcost * discount;
                SummLabel2.Content = $"{totaldetailcostwithdiscount:N2} ₽";
                // Обновите TotalCostLabel
                decimal totalCost = selectedServices.Sum(service => service.Cost * service.Count) + totaldetailcostwithdiscount;
                decimal discountedTotalCost = totalCost;

                TotalCostLabel.Content = $"{discountedTotalCost:N2} ₽";
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (ServicesGrid.SelectedItem != null)
            {
                var serviceToRemove = (Service)ServicesGrid.SelectedItem;

                // Удаление услуги из списка
                selectedServices.Remove(serviceToRemove);

                // Обновление DataGrid для отображения выбранных услуг
                ServicesGrid.ItemsSource = selectedServices;

                // Обновление Label с общей стоимостью
                UpdateTotalCostLabel();
                
            }
        }

        private void ServicesGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            UpdateTotalCostLabel();
            
        }

        private void Button_Click_RemoveDetails(object sender, RoutedEventArgs e)
        {
            if (DetailsGrid.SelectedItem != null)
            {
                var detailToRemove = (Detail)DetailsGrid.SelectedItem;
                selectedDetails.Remove(detailToRemove);
                DetailsGrid.ItemsSource = selectedDetails;
                UpdateTotalCostLabel();
                UpdateTotalCostDetail();
            }
        }
        private void UpdateTotalCost2()
        {
            // Получите значение скидки из DiscountTextBox
            decimal discount = 0;
            if (decimal.TryParse(DiscountTextBox.Text, out discount))
            {
                discount /= 100; // Переведите проценты в десятичное значение

                // Обновите TotalCostLabel
                decimal totalCost = selectedServices.Sum(service => service.TotalCost) + selectedDetails.Sum(detail => detail.TotalCost);
                decimal discountedTotalCost = totalCost - (totalCost * discount);

                TotalCostLabel.Content = $"{discountedTotalCost:N2} ₽";
            }
            // Получите значение скидки из DiscountTextBox
            
            if (decimal.TryParse(DiscountTextBox.Text, out discount))
            {
                discount /= 100; // Переведите проценты в десятичное значение

                // Обновите TotalCostLabel
                decimal totalCost = selectedServices.Sum(service => service.TotalCost) + selectedDetails.Sum(detail => detail.TotalCost);
                decimal discountedTotalCost = totalCost - (totalCost * discount);

                TotalCostLabel.Content = $"{discountedTotalCost:N2} ₽";
            }
        }

        private void Button_Click_AddDetails(object sender, RoutedEventArgs e)
        {
            if (DetailsBox.SelectedItem != null)
            {
                var selectedDetail = (DataRowView)DetailsBox.SelectedItem;
                var DetalName2 = selectedDetail["DetalName"].ToString();
                var costString = selectedDetail["Cost"].ToString();

                var detail = new Detail
                {
                    DetalName = DetalName2,
                    Cost = decimal.Parse(costString, NumberStyles.Currency, CultureInfo.GetCultureInfo("ru-RU")),
                    Quantity = 1 // По умолчанию устанавливаем количество в 1
                };

                // Проверка наличия детали в списке
                if (!selectedDetails.Any(d => d.DetalName == detail.DetalName))
                {
                    detail.OriginalCost = detail.Cost; // Сохранить исходную цену
                    selectedDetails.Add(detail);

                    // Обновление DataGrid для отображения выбранных деталей
                    DetailsGrid.ItemsSource = null;
                    DetailsGrid.ItemsSource = selectedDetails;

                    // Убедитесь, что выбранный элемент также отображается в ComboBox
                    DetailsBox.SelectedItem = detail;
                    UpdateTotalCostDetail();
                    UpdateTotalCostLabel();
                   // RecalculatePrices(); // Вызов пересчета цен после добавления детали
                }
            }
        }

        private void DetailsGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            decimal totalCost = 0;

            // Пересчитываем общую сумму
            foreach (Detail detail in DetailsGrid.SelectedItems)
            {
                totalCost += detail.TotalCost;
            }

            // Обновляем SummLabel2
            SummLabel2.Content = $"{totalCost.ToString("0.00")} ₽.";
           
            UpdateTotalCostLabel();
            UpdateTotalCostDetail();
        }
        private void SelectedDetails_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            decimal totalSum = 0;

            // Пройдитесь по всем деталям и сложите стоимость
            foreach (var detail in selectedDetails)
            {
                totalSum += detail.TotalCost;
            }

            // Обновите SummLabel2
            SummLabel2.Content = $"{totalSum.ToString("N2")} ₽";
            UpdateTotalCostLabel();

        }
        private void DetailsGrid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            if (e.Column.Header.ToString() == "Цена")
            {
                TextBox textBox = e.EditingElement as TextBox;
                if (textBox != null)
                {
                    textBox.Text = ""; // Очистка содержимого ячейки при начале редактирования
                }
            }
        }

        private void DetailsGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                if (e.Column.Header.ToString() == "Количество")
                {
                    var editedDetail = (Detail)e.Row.Item;

                    if (int.TryParse(((TextBox)e.EditingElement).Text, out int newQuantity))
                    {
                        editedDetail.Quantity = newQuantity;
                        UpdateTotalCostLabel(); 
                    }
                }
                else if (e.Column.Header.ToString() == "Цена")
                {
                    var editedDetail = (Detail)e.Row.Item;

                    if (decimal.TryParse(((TextBox)e.EditingElement).Text, out decimal newCost))
                    {
                        editedDetail.Cost = newCost;
                        UpdateTotalCostLabel();
                    }
                }
            }
        }
      
        private void UpdateTotalCostLabel2()
        {
            // Получите значение скидки из DiscountTextBox
            decimal discount = 0;
            if (decimal.TryParse(DiscountTextBox.Text, out discount))
            {
                discount /= 100; // Переведите проценты в десятичное значение

                // Обновите TotalCostLabel
                decimal totalCost = selectedServices.Sum(service => service.Cost * service.Count) + selectedDetails.Sum(detail => detail.TotalCost);
                decimal discountedTotalCost = totalCost - (totalCost * discount);

                TotalCostLabel.Content = $"{discountedTotalCost:N2} ₽";
            }
        }

        private string ReadOrderNumberFromDocument(string documentPath)
        {
            string localDocumentPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ИПЗаказ.docx");
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Open(localDocumentPath);

            // Извлекаем текущий номер из существующей закладки "OrderNumber"
            Word.Range range = doc.Bookmarks["OrderNumber"].Range;
            string orderNumber = range.Text;
            int currentOrderNumber = int.Parse(orderNumber) + 1;

            // Обновляем значение в существующей закладке "OrderNumber" на месте
            range.Text = currentOrderNumber.ToString();

            // Создаем новую закладку "OrderNumber" с обновленным значением
            range.Bookmarks.Add("OrderNumber");

            // Сохраняем документ
            doc.Save();

            // Закрываем и выходим из Word
            doc.Close();
            wordApp.Quit();

            // Возвращаем обновленное значение
            return currentOrderNumber.ToString();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {


            // Определение пути к шаблону и копии
            string templatePath = "ИПЗаказ.docx";
            string copyPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Копия_ИПЗаказ.docx");

            string currentOrderNumber =  OrderNumberBox.Text; // Получаем значение из TextBox

            // Создание копии документа
            File.Copy(templatePath, copyPath, true);

            // Открываем Word
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = true;
            Word.Document doc = wordApp.Documents.Open(copyPath);

            // Заменяем метку на новый номер заказ-наряда
            ReplaceTextInWordDocument(doc, "<НомерЗак>", currentOrderNumber); // Заменяем метку на новое значение


            // Заменяем остальные метки в документе
            ReplaceTextInWordDocument(doc, "<НомТелефонаЗаказчика>", ZakazTele.Text);
            ReplaceTextInWordDocument(doc, "<ФИОЗаказчика>", FIOZakaz.Text);
            ReplaceTextInWordDocument(doc, "<ДатаПрин>", DateTime.Now.ToString("dd.MM.yyyy"));
            ReplaceTextInWordDocument(doc, "<Время>", DateTime.Now.ToString("HH:mm:ss"));
            DateTime selectedDate = DatePickerDeadline.SelectedDate ?? DateTime.Now; // Используйте текущую дату, если ничего не выбрано

            // Преобразуйте дату в строку с нужным форматом
            string formattedDate = selectedDate.ToString("dd.MM.yyyy");

            // Замените текст в документе
            ReplaceTextInWordDocument(doc, "<ДатаИсп>", formattedDate);
            ReplaceTextInWordDocument(doc, "<Марка>", MarksAutos.Text.ToString());
            ReplaceTextInWordDocument(doc, "<Модель>", ModelAutos.SelectedValue?.ToString());
            
            ReplaceTextInWordDocument(doc, "<СкидкаР>", DiscountServBox.Text);
            ReplaceTextInWordDocument(doc, "<СкидкаЗЧ>", DiscountDetailBox.Text);
            ReplaceTextInWordDocument(doc, "<ОбщРаб>", SummLabel.Content.ToString());

            // Форматирование и замена меток с числовыми значениями
            ReplaceTextInWordDocument(doc, "<ОбщС>", TotalCostLabel.Content.ToString());
            ReplaceTextInWordDocument(doc, "<ОбщСЗ>", SummLabel2.Content.ToString());

            ReplaceTextInWordDocument(doc, "<КонтрАгент>", CounterAgentComboBox.Text);
            InsertDataFromDataGridIntoWordTable(doc);
            InsertDataFromDataGrid2IntoWordTable(doc);
            // Генерация имени файла для сохранения 
            string pdfFileName = $"Заказ_{currentOrderNumber}.docx";
            string saveFolderPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Сохраненные_заказы");
            string pdfFilePath = System.IO.Path.Combine(saveFolderPath, pdfFileName);

            if (File.Exists(pdfFilePath))
            {
                MessageBoxResult result = MessageBox.Show("Файл с таким номером заказа уже существует. Вы уверены, что хотите перезаписать заказ?", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.No)
                {
                    // Действия, если пользователь не хочет обнулять заказы
                    doc.Close();
                    wordApp.Quit();
                    
                }
                else
                {
                    // Сохранение документа в формате DOCX
                    doc.SaveAs2(pdfFilePath, Word.WdSaveFormat.wdFormatDocumentDefault);
                    UpdateZakazNar();
                }
            }
            else
            {
                // Сохранение документа в формате DOCX
                doc.SaveAs2(pdfFilePath, Word.WdSaveFormat.wdFormatDocumentDefault);
                UpdateZakazNar();
            }

          



            // Сохранение документа в формате PDF

            // doc.Close();

            // Закрытие Word
            //wordApp.Quit();
            //Process.Start(pdfFilePath);

        }


        private void ReplaceTextInWordDocument(Word.Document doc, string searchText, string replaceText)
        {
            Word.Range range = doc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: searchText, ReplaceWith: replaceText, Replace: Word.WdReplace.wdReplaceAll);
        }
        private void InsertDataFromDataGridIntoWordTable(Word.Document doc)
        {
            // Найти таблицу в Word документе, где вы хотите вставить данные.
            Word.Table table = doc.Tables[3]; // Здесь 1 - номер таблицы в документе, может быть другим

            // Пройдитесь по строкам в DataGrid и вставьте их в таблицу Word.
            for (int i = 0; i < ServicesGrid.Items.Count; i++)
            {
                if (i < table.Rows.Count)
                {
                    var dataGridItem = ServicesGrid.Items[i] as Service;
                    table.Rows[i + 2].Cells[2].Range.Text = dataGridItem.ServiceName; // Вставляем имя работы во 2-й столбец
                    table.Rows[i + 2].Cells[3].Range.Text = dataGridItem.Cost.ToString("N2"); // Вставляем стоимость в 3-й столбец
                    string formattedCount = dataGridItem.Count.ToString("N2").Replace(",", "").TrimEnd('0').TrimEnd('.');
                    table.Rows[i + 2].Cells[4].Range.Text = formattedCount;
                    table.Rows[i + 2].Cells[5].Range.Text = dataGridItem.TotalCost.ToString("N2"); // Вставляем стоимость в 3-й столбец
                }
                else
                {
                    // Если строк в таблице не хватает, добавьте новую строку и вставьте данные.
                    var dataGridItem = ServicesGrid.Items[i] as Service;
                    var newRow = table.Rows.Add();
                    newRow.Cells[2].Range.Text = dataGridItem.ServiceName; // Вставляем имя работы во 2-й столбец
                    newRow.Cells[3].Range.Text = dataGridItem.Cost.ToString("N2"); // Вставляем стоимость в 3-й столбец
                }
            }
        }
        private void InsertDataFromDataGrid2IntoWordTable(Word.Document doc)
        {
            if (doc == null || DetailsGrid == null)
            {
                return;
            }

            // Найти таблицу в Word документе, где вы хотите вставить данные.
            Word.Table table = doc.Tables[4]; // Здесь 2 - номер таблицы в документе, может быть другим

            // Перебрать строки во втором DataGrid и вставить данные в таблицу Word
            for (int i = 0; i < DetailsGrid.Items.Count; i++)
            {
                var dataGridItem = DetailsGrid.Items[i] as Detail; // Замените "YourDataClass" на фактический класс данных во втором DataGrid

                // Вставляем номер строки
                table.Rows[i + 2].Cells[1].Range.Text = (i + 1).ToString();

                // Вставляем наименование из DataGrid во второй столбец таблицы Word
                table.Rows[i + 2].Cells[2].Range.Text = dataGridItem.DetalName;

                // Вставляем количество из DataGrid в третий столбец таблицы Word
                table.Rows[i + 2].Cells[3].Range.Text = dataGridItem.Cost.ToString("N2");

                // Вставляем цену из DataGrid в четвертый столбец таблицы Word
                string formattedCount = dataGridItem.Quantity.ToString().TrimEnd('0');
                table.Rows[i + 2].Cells[4].Range.Text = formattedCount;
                // Вставляем цену из DataGrid в четвертый столбец таблицы Word
                table.Rows[i + 2].Cells[5].Range.Text = dataGridItem.TotalCost.ToString("N2");
            }
        }

        private string FormatNumericValue(object value)
        {
            if (decimal.TryParse(value.ToString(), out decimal numericValue))
            {
                return string.Format("{0:N}", numericValue);
            }

            return value.ToString();
        }

        private void CounterAgent_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CounterAgentComboBox.SelectedItem != null)
            {
                CounterAgent selectedAgent = (CounterAgent)CounterAgentComboBox.SelectedItem;

                // Найти данные выбранного контрагента по его имени в Excel файле
                string excelFilePath = "ServicesList.xlsx"; // Путь к файлу относительно текущей директории
                try
                {
                    using (var workbook = new XLWorkbook(excelFilePath))
                    {
                        var worksheet = workbook.Worksheet("CounterAgents");

                        var agentRow = worksheet.RowsUsed()
                            .FirstOrDefault(row => row.Cell(1).Value.ToString() == selectedAgent.AgentName);

                        if (agentRow != null)
                        {
                            // Получить данные из Excel и поместить их в текстовые поля
                            string agentPhone = agentRow.Cell(2).Value.ToString(); // Предположим, что номер телефона находится во втором столбце (B)

                            // Помещаем данные в текстовые поля
                            FIOZakaz.Text = selectedAgent.AgentName;
                            ZakazTele.Text = agentPhone;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}");
                }
            }
            if (CounterAgentComboBox.SelectedItem == null)
            {
                FIOZakaz.Text = "";
                ZakazTele.Text = "";
            }
        }
        private void CarTypeRadioButton2_Checked(object sender, RoutedEventArgs e)
        {
           
            isCargoCar = true; // Установить флаг "Грузовой" автомобиль
            RecalculatePrices(); // Вызвать пересчет цен
                                 // Проверяем, выбран ли CheckBox
            string excelFilePath = "ServicesList.xlsx"; // Путь к файлу относительно текущей директории
            try
            {
                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    var worksheet = workbook.Worksheet("List1"); // Имя вашего листа

                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add("ServiceName");
                    dt.Columns.Add("Cost", typeof(string));
                    dt.Columns.Add("Count", typeof(string)); // Добавьте столбец "Count" как string
                    dt.Columns.Add("GruzCost", typeof(string));
                    foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропустить первую строку с заголовками
                    {
                        var serviceName = row.Cell(1).Value.ToString();
                        var cost = row.Cell(4).Value.ToString(); // Сохраните стоимость как строку


                        if (decimal.TryParse(cost, out decimal parsedCost))

                        {
                            dt.Rows.Add(serviceName, cost);

                            // Найдите строку "Цена" и сохраните её стоимость и gruzCost
                            if (serviceName == "Цена")
                            {
                                originalServiceCost = parsedCost;
                                // Предположим, что "Цена" в Excel соответствует услуге "Цена"
                                Service priceService = new Service
                                {
                                    ServiceName = "Цена",
                                    Cost = parsedCost,

                                };
                                // Добавляем созданный элемент в ComboBox
                                dt.Rows.Add(priceService.ServiceName, priceService.Cost.ToString(), priceService.GruzCost.ToString());
                            }
                        }
                    }

                    // Заполните ComboBox данными из DataTable
                    Services.ItemsSource = dt.DefaultView;

                    // Вставьте этот код для настройки шаблона элементов ComboBox
                    Services.DisplayMemberPath = null; // Очистите настройку отображаемого пути
                    Services.SelectedValuePath = null; // Очистите настройку значения

                    // Создайте шаблон для элементов ComboBox
                    Services.ItemTemplate = new DataTemplate();

                    var textBlock = new FrameworkElementFactory(typeof(TextBlock));
                    textBlock.SetBinding(TextBlock.TextProperty, new Binding("ServiceName"));
                    textBlock.SetBinding(TextBlock.TagProperty, new Binding("Cost"));


                    Services.ItemTemplate.VisualTree = textBlock;
                    RecalculatePrices(); // Вызвать пересчет цен
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}");
            }
           

           
                try
                {
                    using (var workbook = new XLWorkbook(excelFilePath))
                    {
                        var worksheet = workbook.Worksheet("List2"); // Имя вашего листа

                        System.Data.DataTable dt = new System.Data.DataTable();
                        dt.Columns.Add("DetalName");
                        dt.Columns.Add("Cost", typeof(string)); // Установите тип данных столбца как string
                        dt.Columns.Add("GruzCost", typeof(string)); // Добавьте столбец "GruzCost" как string

                        foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропустить первую строку с заголовками
                        {
                            var detalName = row.Cell(1).Value.ToString();
                            var cost = row.Cell(4).Value.ToString(); // Сохраните стоимость как строку


                            if (decimal.TryParse(cost, out decimal parsedCost))
                            {
                                dt.Rows.Add(detalName, cost);

                                // Найдите строку "Цена" и сохраните её стоимость
                                if (detalName == "Цена")
                                {
                                    originalDetailsCost = parsedCost;
                                }
                            }
                        }

                        // Заполните ComboBox данными из DataTable
                        DetailsBox.ItemsSource = dt.DefaultView;

                        // Вставьте этот код для настройки шаблона элементов ComboBox
                        DetailsBox.DisplayMemberPath = null; // Очистите настройку отображаемого пути
                        DetailsBox.SelectedValuePath = null; // Очистите настройку значения

                        // Создайте шаблон для элементов ComboBox
                        DetailsBox.ItemTemplate = new DataTemplate();

                        var textBlock = new FrameworkElementFactory(typeof(TextBlock));
                        textBlock.SetBinding(TextBlock.TextProperty, new Binding("DetalName"));
                        textBlock.SetBinding(TextBlock.TagProperty, new Binding("Cost"));
                        textBlock.SetBinding(TextBlock.TagProperty, new Binding("GruzCost")); // Добавьте привязку для GruzCost

                        DetailsBox.ItemTemplate.VisualTree = textBlock;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}");
                }
            
        }
       
        private void CarTypeRadioButton1_Checked(object sender, RoutedEventArgs e)
        {
            isCargoCar = false; // Сбросить флаг "Грузовой" автомобиль
           
            string excelFilePath = "ServicesList.xlsx"; // Путь к файлу относительно текущей директории
            try
            {
                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    var worksheet = workbook.Worksheet("List1"); // Имя вашего листа

                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add("ServiceName");
                    dt.Columns.Add("Cost", typeof(string));
                    dt.Columns.Add("Count", typeof(string)); // Добавьте столбец "Count" как string
                    dt.Columns.Add("GruzCost", typeof(string));
                    foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропустить первую строку с заголовками
                    {
                        var serviceName = row.Cell(1).Value.ToString();
                        var cost = row.Cell(3).Value.ToString(); // Сохраните стоимость как строку


                        if (decimal.TryParse(cost, out decimal parsedCost))

                        {
                            dt.Rows.Add(serviceName, cost);

                            // Найдите строку "Цена" и сохраните её стоимость и gruzCost
                            if (serviceName == "Цена")
                            {
                                originalServiceCost = parsedCost;
                                // Предположим, что "Цена" в Excel соответствует услуге "Цена"
                                Service priceService = new Service
                                {
                                    ServiceName = "Цена",
                                    Cost = parsedCost,

                                };
                                // Добавляем созданный элемент в ComboBox
                                dt.Rows.Add(priceService.ServiceName, priceService.Cost.ToString(), priceService.GruzCost.ToString());
                            }
                        }
                    }

                    // Заполните ComboBox данными из DataTable
                    Services.ItemsSource = dt.DefaultView;

                    // Вставьте этот код для настройки шаблона элементов ComboBox
                    Services.DisplayMemberPath = null; // Очистите настройку отображаемого пути
                    Services.SelectedValuePath = null; // Очистите настройку значения

                    // Создайте шаблон для элементов ComboBox
                    Services.ItemTemplate = new DataTemplate();

                    var textBlock = new FrameworkElementFactory(typeof(TextBlock));
                    textBlock.SetBinding(TextBlock.TextProperty, new Binding("ServiceName"));
                    textBlock.SetBinding(TextBlock.TagProperty, new Binding("Cost"));


                    Services.ItemTemplate.VisualTree = textBlock;
                    RecalculatePrices(); // Вызвать пересчет цен
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}");
            }
           

            
                try
                {
                    using (var workbook = new XLWorkbook(excelFilePath))
                    {
                        var worksheet = workbook.Worksheet("List2"); // Имя вашего листа

                        System.Data.DataTable dt = new System.Data.DataTable();
                        dt.Columns.Add("DetalName");
                        dt.Columns.Add("Cost", typeof(string)); // Установите тип данных столбца как string
                        dt.Columns.Add("GruzCost", typeof(string)); // Добавьте столбец "GruzCost" как string

                        foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропустить первую строку с заголовками
                        {
                            var detalName = row.Cell(1).Value.ToString();
                            var cost = row.Cell(3).Value.ToString(); // Сохраните стоимость как строку


                            if (decimal.TryParse(cost, out decimal parsedCost))
                            {
                                dt.Rows.Add(detalName, cost);

                                // Найдите строку "Цена" и сохраните её стоимость
                                if (detalName == "Цена")
                                {
                                    originalDetailsCost = parsedCost;
                                }
                            }
                        }

                        // Заполните ComboBox данными из DataTable
                        DetailsBox.ItemsSource = dt.DefaultView;

                        // Вставьте этот код для настройки шаблона элементов ComboBox
                        DetailsBox.DisplayMemberPath = null; // Очистите настройку отображаемого пути
                        DetailsBox.SelectedValuePath = null; // Очистите настройку значения

                        // Создайте шаблон для элементов ComboBox
                        DetailsBox.ItemTemplate = new DataTemplate();

                        var textBlock = new FrameworkElementFactory(typeof(TextBlock));
                        textBlock.SetBinding(TextBlock.TextProperty, new Binding("DetalName"));
                        textBlock.SetBinding(TextBlock.TagProperty, new Binding("Cost"));
                        textBlock.SetBinding(TextBlock.TagProperty, new Binding("GruzCost")); // Добавьте привязку для GruzCost

                        DetailsBox.ItemTemplate.VisualTree = textBlock;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}");
                }
            
        }
        private void RecalculatePrices()
        {
            foreach (Service service in selectedServices)
            {
                if (isCargoCar)
                {
                    service.Cost = service.OriginalCost * 1; // Увеличить цену на 25%
                }
                else
                {
                    service.Cost = service.OriginalCost; // Вернуть цену к исходной
                }
            }

            foreach (Detail detail in selectedDetails)
            {
                if (isCargoCar)
                {
                    detail.Cost = detail.OriginalCost * 1; // Увеличить цену на 25%
                }
                else
                {
                    detail.Cost = detail.OriginalCost; // Вернуть цену к исходной
                }
            }

            // Обновить DataGrid для отображения измененных цен
            ServicesGrid.Items.Refresh();
            DetailsGrid.Items.Refresh();

            // Обновить общую стоимость
            UpdateTotalCostLabel();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            string templatePath = "ИПЗаказ.docx";
            string copyPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Копия_ИПЗаказ.docx");


            File.Copy(templatePath, copyPath, true);


            // Проверка, открыт ли документ
            Word.Application wordApp;
            Word.Document doc;
            bool isDocOpened = false;

            try
            {
                // Попытка открыть документ
                wordApp = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                doc = wordApp.ActiveDocument;
                isDocOpened = true;
            }
            catch
            {
                // Если документ не открыт, создаем новый экземпляр Word
                wordApp = new Word.Application();
                wordApp.Visible = true;
                doc = wordApp.Documents.Open(copyPath);
            }

            // Если документ открыт, закрываем его и повторно выполняем операции
            if (isDocOpened)
            {
                foreach (Process process in Process.GetProcessesByName("WINWORD"))
                {
                    process.Kill(); // Закрыть все процессы Word
                }

                /*doc.Close();
                wordApp.Quit();*/
                // Повторяем операции, начиная с открытия документа
                Button_Click_2(sender, e);
                return;
            }


            // Заменяем остальные метки в документе
            ReplaceTextInWordDocument(doc, "<НомТелефонаЗаказчика>", ZakazTele.Text);
            ReplaceTextInWordDocument(doc, "<ФИОЗаказчика>", FIOZakaz.Text);
            ReplaceTextInWordDocument(doc, "<ДатаПрин>", DateTime.Now.ToString("dd.MM.yyyy"));
            ReplaceTextInWordDocument(doc, "<Время>", DateTime.Now.ToString("HH:mm:ss"));
            DateTime selectedDate = DatePickerDeadline.SelectedDate ?? DateTime.Now; // Используйте текущую дату, если ничего не выбрано

            // Преобразуйте дату в строку с нужным форматом
            string formattedDate = selectedDate.ToString("dd.MM.yyyy");

            // Замените текст в документе
            ReplaceTextInWordDocument(doc, "<ДатаИсп>", formattedDate);
            ReplaceTextInWordDocument(doc, "<Марка>", MarksAutos.Text.ToString());
            ReplaceTextInWordDocument(doc, "<Модель>", ModelAutos.SelectedValue?.ToString());

            ReplaceTextInWordDocument(doc, "<СкидкаР>", DiscountServBox.Text);
            ReplaceTextInWordDocument(doc, "<СкидкаЗЧ>", DiscountDetailBox.Text);
            ReplaceTextInWordDocument(doc, "<ОбщРаб>", SummLabel.Content.ToString());

            // Форматирование и замена меток с числовыми значениями
            ReplaceTextInWordDocument(doc, "<ОбщС>", TotalCostLabel.Content.ToString());
            ReplaceTextInWordDocument(doc, "<ОбщСЗ>", SummLabel2.Content.ToString());
            if(CounterAgentComboBox.SelectedItem != null)
            {
                ReplaceTextInWordDocument(doc, "<КонтрАгент>", CounterAgentComboBox.Text);
            }
            else
            {
                ReplaceTextInWordDocument(doc, "<КонтрАгент>", FIOZakaz.Text);
            }
            
            InsertDataFromDataGridIntoWordTable(doc);
            InsertDataFromDataGrid2IntoWordTable(doc);
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
           
            string filePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "order_number.txt");
            MessageBoxResult result = MessageBox.Show("Вы уверены, что хотите обнулить заказы?", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    // Проверка наличия файла
                    if (File.Exists(filePath))
                    {
                        // Чтение содержимого файла
                        string[] lines = File.ReadAllLines(filePath);
                        string newOrderNumber = "";

                        // Изменение значения OrderNumber
                        for (int i = 0; i < lines.Length; i++)
                        {
                            if (lines[i].StartsWith("OrderNumber="))
                            {
                                lines[i] = "OrderNumber=1"; // Установка нового значения OrderNumber
                                newOrderNumber = "1"; // Сохраняем новое значение OrderNumber
                                break; // Выход из цикла после изменения
                            }
                        }

                        // Запись измененных данных обратно в файл
                        File.WriteAllLines(filePath, lines);

                        // Присваиваем новое значение OrderNumberBox
                        OrderNumberBox.Text = newOrderNumber;
                    }
                    else
                    {
                        // Если файл не найден, можно вывести сообщение об ошибке или создать новый файл с нужным содержимым
                        // Здесь можно добавить логику для создания файла с OrderNumber=1
                        File.WriteAllText(filePath, "OrderNumber=1");
                        OrderNumberBox.Text = "1"; // Присваиваем начальное значение OrderNumberBox
                    }
                }
                catch (Exception ex)
                {
                    // Обработка ошибок при записи в файл или чтении OrderNumberBox
                    Console.WriteLine("Ошибка: " + ex.Message);
                }
            }
            else
            {

            }
                
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            selectedServices.Clear();
            selectedDetails.Clear();
        }

        private void ServicesGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void ServicesGrid_KeyUp(object sender, KeyEventArgs e)
        {
            
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            CounterAgentComboBox.SelectedItem = null;
            FIOZakaz.Text = "";
            ZakazTele.Text = "";
        }
    }
}
