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
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Window = System.Windows.Window;


namespace ЧЕК
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public double Count = 0;//Стоимость тарифа за минуту
        public MainWindow()
        {
            InitializeComponent();
            Tariff.Items.Add("Тариф 1");// 0
            Tariff.Items.Add("Тариф 2");// 1
        }

        private void CreateCheck_Click(object sender, RoutedEventArgs e)
        {
            if (Tariff.SelectedIndex == 0)
            {
                //MessageBox.Show("Тариф 1");
                Create_Check(Count = 0.5);
            }
            else if (Tariff.SelectedIndex == 1) 
            {
                //MessageBox.Show("Тариф 2");
                Create_Check(Count = 1);
            }
        }

        private void Create_Check(double Count)
        {
            var fileName = $"{"Чек"}_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.docx";//Имя + указание даты создания файла воизбежании замены файла с таким же названием
            var savePath = System.IO.Path.GetFullPath($@"..\..\..\WORD\{fileName}");//Получение абсолютного пути для ворд файла
            //var savePath = @"\WORD\" + fileName;

            var wordApp = new Application();
            var document = wordApp.Documents.Add();
            document.Content.SetRange(0, 0);
            var companyName = "ООО Читатель";
            var welcomeText = "Добро пожаловать";
            var kkmNumber = "ККМ 00075411 #3969";
            var inn = "ИНН 1087746942040";
            var ekls = "ЭКЛЗ 3851495566";
            Random random = new Random();
            int num = random.Next(000000001, 999999999);
            var checkNumber = $"Чек №{num}";
            var dateTime = $"{DateTime.Now.ToString("yyyyMMdd_HHmmss")} СИС.";
            var line = "----------------------";
            document.Content.Text = $"{companyName}" +
                $"\n{welcomeText}" +
                $"\n{kkmNumber}" +
                $"\n{inn}" +
                $"\n{ekls}" +
                $"\n{checkNumber}" +
                $"\n{dateTime}" +
                $"\n{line}" +
                $"\nТовар: " +
                $"\nНазвание тарифа: Тариф 1\n " +
                $"\nЦена: {Count}" +
                $"\n{line}";
            document.SaveAs2(savePath);
            document.Close();
            wordApp.Quit();
            MessageBox.Show($"Чек {fileName} создан и расположен в: {savePath}");
        }
    }
}
