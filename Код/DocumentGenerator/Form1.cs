using System;
using System.Windows.Forms;

namespace DocumentGenerator
{
    //// Класс программы
    public partial class Program1 : Form
    {
        #region Поля Класса
        public static int MaxDocNumber { get; set; }
        public static int DocNumber { get; set; }
        public static string SaveAdress { get; set; }
        public static string AdressData { get; set; }
        public static string AdressTemplates { get; set; }
        public static Data_Form Data { get; set; }
        #endregion

        #region Конструктор Класса
        public Program1()
        {
            InitializeComponent();
            DocNumber = 1;
            MaxDocNumber = 5;
            progressBar1.Maximum = MaxDocNumber;
        }
        #endregion

        #region Методы Класса
        // Метод выбора Excel опросника
        public void ChooseData()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            ofd.Title = "Выберите документ Excel";

            if (ofd.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали файл для открытия", "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            AdressData = ofd.FileName;
        }
        // Метод чтения Excel опросника
        public void SetData(Data_Form Data)
        {
            Excel_Document ExOpen = new Excel_Document(Data);
            Start:
            try
            {
                ExOpen.Open(AdressData);
                ExOpen.CopyData();
            }
            catch 
            {
                MessageBox.Show("Выбранный бланк не соответсвует формату", "Ошибка");
                ExOpen.App.Quit();
                ExOpen.MemoryClear();
                ChooseData();
                goto Start;
            }
        }

        // Выбор Excel опросника
        private void СhooseData_Click(object sender, EventArgs e)
        {
            ChooseData();
        }

        //Выбор папки с шаблонами
        private void ChooseFolderTemplates_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                AdressTemplates = folderBrowserDialog1.SelectedPath;
            }
        }

        //Выбор места для сохранения
        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                SaveAdress = folderBrowserDialog1.SelectedPath;
            }
        }

        //Генерация актов
        private void Generate_Button_Click(object sender, EventArgs e)
        {
            // Проверка всех путей
            if (SaveAdress == null || AdressData == null || AdressTemplates == null)
            {
                MessageBox.Show("Не были выбраны все пути", "Ошибка");
                return;
            }

            progressBar1.Visible = true;
            progressBar1.Value = 0;

            //Копирование данных
            Data_Form Data = new Data_Form();
            SetData(Data);


            //Создание документа Excel
            Excel_Document Document = new Excel_Document(Data);

                for (byte i = 1; i <= 5; i++)
                {
                    try
                    {
                        Document.Create(i);
                    }
                    catch 
                    {
                        MessageBox.Show($"Отсутствует шаблон {(Excel_Document.Templates)i}", "На один акт меньше");

                        Document.App.Quit();
                        Document.MemoryClear();
                    }
                    progressBar1.Value++;
                }
        }
        #endregion
    }
}
