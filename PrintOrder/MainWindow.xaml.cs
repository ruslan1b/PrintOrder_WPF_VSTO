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
using System.IO;


using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Globalization;
using MaterialDesignThemes.Wpf;
using MaterialDesignColors;
using LinqToDB;
using System.Data.Linq;
using MySql.Data.MySqlClient;
using DataModels;
using LinqToDB.DataProvider.SqlServer;

using BarcodeLib;
using System.Drawing;
using System.Drawing.Imaging;
using Gma.QrCodeNet.Encoding;
using Gma.QrCodeNet.Encoding.Windows.Render;

namespace PrintOrder
{

    public partial class MainWindow : System.Windows.Window
    {
        Excel._Application application;
        //Excel.__Document document;
        Excel.Workbooks workbooks;
        Excel.Workbook workbook;
        Excel.Sheets sheets;
        Excel.Worksheet worksheet;
        Excel.Range cells;

        List<string> selectionChanged = new List<string>();
        public string supplierSelect, recipientSelect, payerSelect, carSelect, orderSelect, dateSelect, worksSelect, goodsSelect, worksSelect_Copy, goodsSelect_Copy;
        string Country, Index, City, Street, NumberOffice, Tel1, Tel2, Email, Account, Bank, MFO, EDRPOU, CodeInd, NumberSvid;
        string Address, CityR, Tel, EmailR, IndexR, Mob;
        string AddressPay, CityPay, TelPay, EmailPay, IndexPay, MobPay;
        string NumBody, TypeIngine, Year, NumIngine, NumCountry, Agregat, VIN;

        

        string UnitOfMeasureW, AmountW, PriceW, SumW;
        string UnitOfMeasureG, AmountG, PriceG, SumG;
        int numberGoods, numberWorks;
        string getQrCode, getBarCode;

        Object missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;

        ComboBox ComboBoxWorks_Copy = new ComboBox();

        public MainWindow()
        {
            InitializeComponent();
        }
        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
            // Initialize an OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Set filter and RestoreDirectory
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Filter = "Excel documents(*.xls;*.xlsx)|*.xls;*.xlsx";

            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                if (openFileDialog.FileName.Length > 0)
                {
                    txbSelectedExcelFile.Text = openFileDialog.FileName;
                }
            }
        }
        private void btnConvertToForm_Click(object sender, RoutedEventArgs e)
        {
            if (txbSelectedExcelFile.Text != null)
            {
                MessageBox.Show("Файл знайдено");
            }
            else MessageBox.Show("Файл не знайдено");

            Test(txbSelectedExcelFile.Text);
        }
        private void Test(string excelFilename)
        {
            //создаем обьект приложения excel
            application = new Excel.Application();
            // создаем путь к файлу
            Object templatePathObj = txbSelectedExcelFile.Text;

            // если вылетим не этом этапе, приложение останется открытым
            try
            {
                workbook = application.Workbooks.Open(excelFilename);

            }
            catch (Exception error)
            {
                MessageBox.Show("Произошла ошибка");
                application.Quit();
                application = null;
                throw error;
            }
            ConvertToForm();
            application.Visible = true;
        }
        private void ConvertToForm()
        {
            int m, n;       

            workbook = application.Workbooks.Open(txbSelectedExcelFile.Text,
                              Type.Missing, Type.Missing, Type.Missing,
            "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

            //workbook = workbooks[1];
            workbook.Activate();
            //Получаем массив ссылок на листы выбранной книги
            sheets = workbook.Worksheets;
            //Получаем ссылку на лист 1
            worksheet = (Excel.Worksheet)sheets.get_Item(1);
            //Выбираем ячейку для вывода A1
            cells = worksheet.get_Range("A1", "A1");

            string content = Convert.ToString(cells.Value2);

            List<string> selectionChanged = new List<string>();

            //виловлення помилки нульового значення дати
            if (Date_Picker.SelectedDate.Value == null)
                MessageBox.Show("Выберите дату");
            dateSelect = Date_Picker.SelectedDate.Value.ToString();

            Dictionary<string, object> nameKeys = new Dictionary<string, object> {
                { "[Получ.Название]" , recipientSelect },
                { "[Платил.Название]" , payerSelect },
                { "[Постав.Название]" , supplierSelect },
                { "[Авто.Марка]" , carSelect },
                { "[НЗ.№]" , orderSelect },
                { "[НЗ.Дата]" , dateSelect },
                { "[Таблица.Нз.Работ.Наименование]" , worksSelect },
                { "[Таблица.Нз.Товар.Наименование]" , goodsSelect },               
            };

            Dictionary<string, object> nameKeys_Copy = new Dictionary<string, object> {
                { "[Таблица.Нз.Работ.Наименование]" , worksSelect_Copy },
                { "[Таблица.Нз.Товар.Наименование]" , goodsSelect_Copy },
            };

            GetDataTableSupplier();
            GetDataTableRecipient();
            GetDataTablePayer();
            GetDataTableDataCar();
            GetDataTableWorks();
            GetDataTableGoods();
            GetDataTableWorks_Copy();
            GetDataTableGoods_Copy();

            Dictionary<string, object> suppliesKeys = new Dictionary<string, object>
            {
                {"[Постав.Страна]" , Country },
                {"[Постав.Индекс]" , Index },
                {"[Постав.Город]" , City },
                {"[Постав.Улица]" , Street },
                {"[Постав.Офис№]" , NumberOffice },
                {"[Постав.Тел1]" , Tel1 },
                {"[Постав.Тел2]" , Tel2 },
                {"[Постав.Email]" , Email },
                {"[Постав.№Р/р]" , Account },
                {"[Постав.НазвБанка]" , Bank },
                {"[Постав.МФО]" , MFO },
                {"[Постав.ЄДРПОУ]" , EDRPOU },
                {"[Постав.ИДН]" , CodeInd },
                {"[Постав.№Свид]" , NumberSvid },

            };
            Dictionary<string, object> recipientKeys = new Dictionary<string, object>
            {
                { "[Получ.Адрес]", Address },
                { "[Получ.Город]", CityR},
                { "[Получ.№Тел]", Tel},
                { "[Получ.Email]", EmailR},
                { "[Получ.Индекс]", IndexR},
                { "[Получ.Моб]", Mob},
            };
            Dictionary<string, object> payerKeys = new Dictionary<string, object>
            {
                { "[Платил.Адрес]", AddressPay },
                { "[Платил.Город]", CityPay},
                { "[Платил.№Тел]", TelPay},
                { "[Платил.Email]", EmailPay},
                { "[Платил.Индекс]", IndexPay},
                { "[Платил.Моб]", MobPay},
            };
            Dictionary<string, object> dataCartKeys = new Dictionary<string, object>
            {
                { "[Авто.№Кузова]", NumBody },
                { "[Авто.ТипДвиг]", TypeIngine},
                { "[Авто.ГодВыпуска]", Year},
                { "[Авто.№Двиг]", NumIngine},
                { "[Авто.НазвАгрегата]", NumCountry},
                { "[Авто.№Гос]", Agregat},
                { "[Авто.VINКод]", VIN},
            };
            Dictionary<string, object> worksKeys = new Dictionary<string, object>
            {
                { "[Таблица.Нз.Работ.№]", numberWorks },
                { "[Таблица.Нз.Работ.Ед]", UnitOfMeasureW },
                { "[Таблица.Нз.Работ.Кол]", AmountW},
                { "[Таблица.Нз.Работ.Цена]", PriceW},
                { "[Таблица.Нз.Работ.Сумма]", SumW},
            };
            Dictionary<string, object> goodsKeys = new Dictionary<string, object>
            {
                { "[Таблица.Нз.Товар.№]", numberGoods },
                { "[Таблица.Нз.Товар.Ед]", UnitOfMeasureG },
                { "[Таблица.Нз.Товар.Кол]", AmountG},
                { "[Таблица.Нз.Товар.Цена]", PriceG},
                { "[Таблица.Нз.Товар.Сумма]", SumG},
            };
            Dictionary<string, object> codeKeysBar = new Dictionary<string, object>
            {
                {"[Постав.ШтрихКод]", getQrCode},                
            };
            Dictionary<string, object> codeKeysQr = new Dictionary<string, object>
            {
                {"[Постав.QRКод]",  getBarCode},
            };

            for (m = 1; m < 120; m++)
            {
                for (n = 1; n < 25; n++)
                {
                    cells = (Excel.Range)worksheet.Cells[m, n];                   
                    if (cells.Value2 == null)
                        continue;

                    string cellValue = cells.Value2.ToString();

                    nameKeys.Keys.ToList().ForEach(key =>
                    {
                        if (cellValue.Contains(key))
                        {
                            cells.Value2 = nameKeys[key] as object;
                        }
                    });

                    nameKeys_Copy.Keys.ToList().ForEach(key =>
                    {
                        if (cellValue.Contains(key))
                        {
                            string cellAdres;
                            char[] adr = new char[5];
                            char[] adr1 = new char[5];

                            int ch;
                            cellAdres = cells.Address.ToString();
                            Console.WriteLine(cellAdres);
                            adr = cellAdres.ToCharArray(1, 1);
                            adr1 = cellAdres.ToCharArray(3, 2);
                            Console.WriteLine(adr);
                            Console.WriteLine(adr1);
                            ch = Int32.Parse((adr1[0].ToString()) + adr1[1]);
                            Console.WriteLine(ch);

                            cells = (Excel.Range)worksheet.Cells[adr[0], ch + 1];
                            Console.WriteLine(cellAdres);
                            cells.Value2 = nameKeys_Copy[key] as object;
                        }
                    });

                    suppliesKeys.Keys.ToList().ForEach(key =>
                    {
                        if (cellValue.Contains(key))
                        {
                            cells.Value2 = suppliesKeys[key] as object;
                        }
                    });
                    recipientKeys.Keys.ToList().ForEach(key =>
                    {
                        if (cellValue.Contains(key))
                        {
                            cells.Value2 = recipientKeys[key] as object;
                        }
                    });
                    payerKeys.Keys.ToList().ForEach(key =>
                    {
                        if (cellValue.Contains(key))
                        {
                            cells.Value2 = payerKeys[key] as object;
                        }
                    });
                    dataCartKeys.Keys.ToList().ForEach(key =>
                    {
                        if (cellValue.Contains(key))
                        {
                            cells.Value2 = dataCartKeys[key] as object;
                        }
                    });


                    worksKeys.Keys.ToList().ForEach(key =>
                    {
                        if (cellValue.Contains(key))
                        {
                            cells.Value2 = worksKeys[key] as object;
                            
                        }
                    });


                    goodsKeys.Keys.ToList().ForEach(key =>
                    {
                        if (cellValue.Contains(key))
                        {
                            cells.Value2 = goodsKeys[key] as object;
                            
                        }
                    });


                    codeKeysBar.Keys.ToList().ForEach(key =>
                    {
                        if (cellValue.Contains(key))
                        {
                            cells.Value2 = codeKeysBar[key] as object;
                            GetBarCode();
                        }
                    });
                    codeKeysQr.Keys.ToList().ForEach(key =>
                    {
                        if (cellValue.Contains(key))
                        {
                            cells.Value2 = codeKeysQr[key] as object;
                            GetQrCode();
                        }
                    });
                    
                }
            }

        }
        private void ComboBoxRecipient_Loaded(object sender, RoutedEventArgs e)
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Recipients
                    select c;

                List<string> data = new List<string>();
                data.Add("Нужно выбрать получателя услуг");

                foreach (var c in q)
                    data.Add(c.Recip.ToString());

                for (int i = 0; i < data.Count; i++)
                {
                    ComboBoxRecipient.Items.Add(data[i]);
                    ComboBoxRecipient.SelectedIndex = 0;
                }
            }

        }
        private void ComboBoxCar_Loaded(object sender, RoutedEventArgs e)
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Datacars
                    select c;

                List<string> data = new List<string>();
                data.Add("Нужно выбрать автомобиль");

                foreach (var c in q)
                    data.Add(c.Car.ToString());

                for (int i = 0; i < data.Count; i++)
                {
                    ComboBoxCar.Items.Add(data[i]);
                    ComboBoxCar.SelectedIndex = 0;
                }
            }
        }
        private void ComboBoxPayer_Loaded(object sender, RoutedEventArgs e)
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Payers
                    select c;

                List<string> data = new List<string>();
                data.Add("Нужно выбрать платильщика услуг");

                foreach (var c in q)
                    data.Add(c.NameOfPayers.ToString());

                for (int i = 0; i < data.Count; i++)
                {
                    ComboBoxPayer.Items.Add(data[i]);
                    ComboBoxPayer.SelectedIndex = 0;
                }
            }
        }
        private void ComboBoxSupplier_Loaded(object sender, RoutedEventArgs e)
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Suppliers
                    select c;

                List<string> data = new List<string>();
                data.Add("Нужно выбрать поставщика услуг");

                foreach (var c in q)
                    data.Add(c.NameOfOrganiz.ToString());

                for (int i = 0; i < data.Count; i++)
                {
                    ComboBoxSupplier.Items.Add(data[i]);
                    ComboBoxSupplier.SelectedIndex = 0;
                }
            }
        }
        private void TextBoxOrder_Loaded(object sender, RoutedEventArgs e)
        {
            //TextBoxOrder.Text = "Введите номер документа";
        }
        private void ComboBoxWorks_Loaded(object sender, RoutedEventArgs e)
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Works
                    select c;

                List<string> data = new List<string>();
                data.Add("Нужно выбрать работы");

                foreach (var c in q)
                    data.Add(c.NameOfWorks.ToString());

                for (int i = 0; i < data.Count; i++)
                {
                    ComboBoxWorks.Items.Add(data[i]);
                    ComboBoxWorks.SelectedIndex = 0;
                }
            }
        }        
        private void ComboBoxGoods_Loaded(object sender, RoutedEventArgs e)
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Goods
                    select c;

                List<string> data = new List<string>();
                data.Add("Нужно выбрать материалы");

                foreach (var c in q)
                    data.Add(c.NameOfGoods.ToString());

                for (int i = 0; i < data.Count; i++)
                {
                    ComboBoxGoods.Items.Add(data[i]);
                    ComboBoxGoods.SelectedIndex = 0;
                }
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            orderSelect = TextBoxOrder.Text.ToString();
        }

        public void ComboBoxRecipient_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedcomboitem = sender as ComboBox;
            string name = selectedcomboitem.SelectedItem as string;
            recipientSelect = name;
        }
        private void ComboBoxPayer_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedcomboitem = sender as ComboBox;
            string name = selectedcomboitem.SelectedItem as string;
            payerSelect = name;
        }
        private void ComboBoxSupplier_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedcomboitem = sender as ComboBox;
            string name = selectedcomboitem.SelectedItem as string;
            supplierSelect = name;
        }
        private void ComboBoxCar_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedcomboitem = sender as ComboBox;
            string name = selectedcomboitem.SelectedItem as string;
            carSelect = name;
        }
        private void ComboBoxWorks_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedcomboitem = sender as ComboBox;
            string name = selectedcomboitem.SelectedItem as string;
            worksSelect = name;
        }
        private void ComboBoxGoods_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedcomboitem = sender as ComboBox;
            string name = selectedcomboitem.SelectedItem as string;
            goodsSelect = name;
        }

        private void GetDataTableSupplier()
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Suppliers
                    where c.NameOfOrganiz == supplierSelect
                    select c;

                foreach (var c in q)
                {
                    Country = c.Country;
                    Index = c.Index;
                    City = c.City;
                    Street = c.Street;
                    NumberOffice = c.NumberOffice;
                    Tel1 = c.Tel1;
                    Tel2 = c.Tel2;
                    Email = c.Email;
                    Account = c.Account;
                    Bank = c.Bank;
                    MFO = c.MFO;
                    EDRPOU = c.EDRPOU;
                    CodeInd = c.CodeInd;
                    NumberSvid = c.NumberSvid;
                }
            }
        }
        private void GetDataTableRecipient()
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Recipients
                    where c.Recip == recipientSelect
                    select c;

                foreach (var c in q)
                {
                    Address = c.Adress;
                    CityR = c.City;
                    Tel = c.Tel;
                    EmailR = c.Email;
                    IndexR = c.Index;
                    Mob = c.Mob;
                }
            }
        }
        private void GetDataTablePayer()
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Payers
                    where c.NameOfPayers == payerSelect
                    select c;

                foreach (var c in q)
                {
                    AddressPay = c.Address;
                    CityPay = c.City;
                    TelPay = c.Tel;
                    EmailPay = c.Email;
                    IndexPay = c.Index;
                    MobPay = c.Mob;
                }
            }
        }
        private void GetDataTableDataCar()
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Datacars
                    where c.Car == carSelect
                    select c;

                foreach (var c in q)
                {
                    NumBody = c.NumBody;
                    TypeIngine = c.TypeIngine;
                    Year = c.Year.ToString();
                    NumIngine = c.NumIngine;
                    NumCountry = c.NumCountry;
                    Agregat = c.Agregat;
                    VIN = c.VIN.ToString();
                }
            }
        }
        private void GetDataTableWorks()
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Works
                    where c.NameOfWorks == worksSelect
                    select c;

                foreach (var c in q)
                {
                    numberWorks = 1;
                    UnitOfMeasureW = c.UnitOfMeasure;
                    AmountW = c.Amount;
                    PriceW = c.Price;
                    SumW = c.Price;
                }
            }
        }
        private void GetDataTableGoods()
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Goods
                    where c.NameOfGoods == goodsSelect
                    select c;

                foreach (var c in q)
                {
                    numberGoods = 1;
                    UnitOfMeasureG = c.UnitOfMeasure;
                    AmountG = c.Amount;
                    PriceG = c.Price;
                    SumG = c.Sum;

                }
            }
        }

        private void GetDataTableWorks_Copy()
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Works
                    where c.NameOfWorks == worksSelect_Copy
                    select c;

                foreach (var c in q)
                {
                    numberWorks = 1;
                    UnitOfMeasureW = c.UnitOfMeasure;
                    AmountW = c.Amount;
                    PriceW = c.Price;
                    SumW = c.Price;
                }
            }
        }
        private void GetDataTableGoods_Copy()
        {
            using (var db = new AutodbDB())
            {
                var q =
                    from c in db.Goods
                    where c.NameOfGoods == goodsSelect_Copy
                    select c;

                foreach (var c in q)
                {
                    numberGoods = 1;
                    UnitOfMeasureG = c.UnitOfMeasure;
                    AmountG = c.Amount;
                    PriceG = c.Price;
                    SumG = c.Sum;

                }
            }
        }

        private void GetBarCode()
        {
            String filePicBar = @"c:\temp\BarCode.png";
            int leftRange = 0, topRange = 0, heightRange = 0;
            
            BarcodeLib.Barcode barCode = new BarcodeLib.Barcode();
            // Create image.
            //System.Drawing.Image img = barCode.Encode(BarcodeLib.TYPE.UPCA, carSelect, System.Drawing.Color.Black, System.Drawing.Color.White, 300, 150);
            System.Drawing.Image img = barCode.Encode(BarcodeLib.TYPE.UPCA, "038000356216", System.Drawing.Color.Black, System.Drawing.Color.White, 300, 150);
            // Create Point for upper-left corner of image.
            //System.Drawing.Point ulCorner = new System.Drawing.Point(100, 100);
            // Draw image to screen.

            //barCode.ImageFormat = System.Drawing.Imaging.ImageFormat.Png;
            // save barcode image into your file system
            barCode.EncodedImage.Save(@"c:\temp\BarCode.png");

            leftRange = (int)cells.Left;
            topRange = (int)cells.Top;
            heightRange = (int)cells.Height;
                      

            worksheet.Shapes.AddPicture(filePicBar, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, leftRange, topRange, 145, heightRange);
        } 

        private void GetQrCode()
        {
            String filePicQr = @"c:\temp\QrCode.png";
           
            int leftRange = 0, topRange = 0;


            QrEncoder qrEncoder = new QrEncoder(ErrorCorrectionLevel.H);
            QrCode qrCode = qrEncoder.Encode(carSelect);

            GraphicsRenderer renderer = new GraphicsRenderer(new FixedModuleSize(5, QuietZoneModules.Two), System.Drawing.Brushes.Black, System.Drawing.Brushes.White);
            using (FileStream stream = new FileStream(@"c:\temp\QrCode.png", FileMode.Create))
            {
                renderer.WriteToStream(qrCode.Matrix, ImageFormat.Png, stream);
            }
            leftRange = (int)cells.Left;
            topRange = (int)cells.Top;
                   
            worksheet.Shapes.AddPicture(filePicQr, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, leftRange, topRange, 145, 145);
        }

        private void CheckBoxWorks_Checked(object sender, RoutedEventArgs e)
        {
           
            

            if (CheckBoxWorks.IsChecked == true)
            {
                ComboBoxItemRecipient.Children.Add(ComboBoxWorks_Copy);
                Grid.SetColumn(ComboBoxWorks_Copy, 1);
                Grid.SetRow(ComboBoxWorks_Copy, 1);

                using (var db = new AutodbDB())
                {
                    var q =
                        from c in db.Works
                        select c;

                    List<string> data = new List<string>();
                    data.Add("Нужно выбрать работы");

                    foreach (var c in q)
                        data.Add(c.NameOfWorks.ToString());

                    for (int i = 0; i < data.Count; i++)
                    {
                        ComboBoxWorks_Copy.Items.Add(data[i]);
                        ComboBoxWorks_Copy.SelectedIndex = 0;
                    }
                }

                
                string name = ComboBoxWorks_Copy.SelectedItem as string;
                worksSelect_Copy = name;
                
            }
            else
            {
                ComboBoxItemRecipient.Children.Remove(ComboBoxWorks_Copy);

            }
            
        }
     
        private void CheckBoxGoods_Checked(object sender, RoutedEventArgs e)
        {
            ComboBox CheckBoxGoods_Copy = new ComboBox();

            if (CheckBoxGoods.IsChecked == true)
            {

                ComboBoxItemRecipient.Children.Add(CheckBoxGoods_Copy);
                Grid.SetColumn(CheckBoxGoods_Copy, 1);
                Grid.SetRow(CheckBoxGoods_Copy, 3);

                using (var db = new AutodbDB())
                {
                    var q =
                        from c in db.Goods
                        select c;

                    List<string> data = new List<string>();
                    data.Add("Нужно выбрать материалы");

                    foreach (var c in q)
                        data.Add(c.NameOfGoods.ToString());

                    for (int i = 0; i < data.Count; i++)
                    {
                        CheckBoxGoods_Copy.Items.Add(data[i]);
                        CheckBoxGoods_Copy.SelectedIndex = 0;
                    }
                }                

            }
            else if (CheckBoxGoods.IsChecked == false)
            {

                ComboBoxItemRecipient.Children.Remove(CheckBoxGoods_Copy);
            }
        }        
    }
}