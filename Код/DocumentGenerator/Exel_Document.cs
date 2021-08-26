using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace DocumentGenerator
{
    //Класс для генерации создания Актов
    class Excel_Document
    {
        #region Поля Класса  
        private string TemplatesАdress { get; set; } // Полный путь к шаблону
        public Application App { get; set; } // Объект Excel
        private Workbooks WorkBooks { get; set; } // Переменная для коллекции рабочих книг
        private Workbook WorkBookTemplates { get; set; } // Переменная для одной рабочей книги
        private Sheets Sheets_ { get; set; } // Переменная для листов
        private Worksheet WSheet { get; set; }  // Переменная для листа
        private Range Cell { get; set; }    // Переменная для ячейки 
        private Data_Form DataForm { get; set; }  // Данные с шаблона
        #endregion

        #region Конструктор Класса
        public Excel_Document(Data_Form DataForm)
        {
            this.DataForm = DataForm;
        }
        #endregion

        #region Методы Класса
        public void Create(byte Template)
        {

            TemplatesАdress = $@"{Program1.AdressTemplates}\{(Templates)Template}.xlsx";
            Open(TemplatesАdress);

            PasteData(Template);

            WorkBookTemplates.SaveAs($@"{Program1.SaveAdress}\{Program1.DocNumber} {WorkBookTemplates.Name}");

            App.Quit(); //  Закрытие окна Excel

            MemoryClear();
        }
        private void DeleteRows(Range CellDel)
        {
            WSheet.Rows[CellDel.Row + 1].Delete();
            WSheet.Rows[CellDel.Row].Delete();
        }
        private void DuplicateRows(Range CellDup)
        {
            CellDup = WSheet.Cells[CellDup.Row + 2, CellDup.Column];
            int I = CellDup.Row - 2;
            CellDup.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, false);
            CellDup.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, false);
            WSheet.Rows[I].Copy(WSheet.Rows[I+2]);
            WSheet.Rows[I+1].Copy(WSheet.Rows[I+3]);
        }
        private void TextWrapping(string Text, Range CellStart, byte TextLengthMaxFirst = 65, byte TextLengthMaxFinal = 100)
        {
            byte IndexReplace = 0;
            string TextRep;
            char Space = Char.Parse(" ");
            byte RowTrigger = 0;

            for (byte i = 0; i < Text.Length; i++)
            {
                if (Text.Length <= TextLengthMaxFirst)
                {
                    if (RowTrigger == 0)
                    {
                        CellStart.MergeArea.Value = Text;
                        CellStart = WSheet.Cells[CellStart.Row + 2, CellStart.Column];
                        DeleteRows(CellStart);
                        break;
                    }

                    else
                    {
                        CellStart.MergeArea.Value = Text;
                        break;
                    }
                }

                if ((Text[i] == Space|| i+1 == Text.Length) && (i >= TextLengthMaxFirst)) //Возможно тут есть косяк во втором условии
                {
                    if (RowTrigger > 0)
                    {
                        DuplicateRows(CellStart);
                    }
                    TextRep = Text.Substring(0, IndexReplace);
                    CellStart.MergeArea.Value = TextRep;
                    Text = Text.Substring(IndexReplace);
                    TextLengthMaxFirst = TextLengthMaxFinal;
                    IndexReplace = 0;
                    CellStart = WSheet.Cells[CellStart.Row + 2, CellStart.Column];
                    RowTrigger++;
                    i = 0;
                    continue;
                }

                if (Text[i] == Space)
                {
                    IndexReplace = i;
                }
            }
        }
        public void Open(string Adress)
        {
            App = new Application();
            WorkBooks = App.Workbooks;
            WorkBookTemplates = WorkBooks.Open(Adress);
            Sheets_ = WorkBookTemplates.Worksheets;
            WSheet = Sheets_.Item[1];
        }
        public void MemoryClear()
        {
            //Освобождение места в памяти
            try
            {
                Marshal.ReleaseComObject(Cell);
                Marshal.ReleaseComObject(WSheet);
                Marshal.ReleaseComObject(Sheets_);
                Marshal.ReleaseComObject(WorkBookTemplates);
                Marshal.ReleaseComObject(WorkBooks);
                Marshal.ReleaseComObject(App);
            }
            catch 
            { 

            }
            finally
            {

                GC.Collect();
            }
        }
        public void CDH(byte NameField)
        {
            CellData cellData = (CellData)NameField;
            string adress = cellData.ToString();
            Cell = WSheet.Cells.Find(adress);

            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
        }
        public void CopyData()
        {
                CDH(0);
                DataForm.BuildObject = Cell.Value;
                CDH(1);
                DataForm.StartDate = DateTime.Parse(Cell.Value);
                CDH(2);
                DataForm.EndDate = DateTime.Parse(Cell.Value);
                CDH(3);
                DataForm.ProjectNumber = Cell.Value;

                Cell = WSheet.Cells.Find(((CellData)19).ToString());
                Cell = WSheet.Cells[Cell.Row + 1, Cell.Column];
                DataForm.ProjectCompany = Cell.Value;

                CDH(20);
                DataForm.BuildingName = Cell.Value;
                CopyCompanyВременноеРешение();

                App.Quit();
                MemoryClear();
        }
        public void CopyCompanyВременноеРешение()
        {
            CDH(4); // 1 Компания
            DataForm.Organization1.CompanyName = Cell.Value;

            Cell = WSheet.Cells[Cell.Row + 1, Cell.Column];
            DataForm.Organization1.CompanyEmployee1.EmployeeName = Cell.Value;
            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
            DataForm.Organization1.CompanyEmployee1.EmployeePost = Cell.Value;
            Cell = WSheet.Cells[Cell.Row + 1, Cell.Column - 1];

            DataForm.Organization1.CompanyEmployee2.EmployeeName = Cell.Value;
            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
            DataForm.Organization1.CompanyEmployee2.EmployeePost = Cell.Value;
            Cell = WSheet.Cells[Cell.Row + 1, Cell.Column - 1];

            DataForm.Organization1.CompanyEmployee3.EmployeeName = Cell.Value;
            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
            DataForm.Organization1.CompanyEmployee3.EmployeePost = Cell.Value;

            CDH(8); // 2 Компания
            DataForm.Organization2.CompanyName = Cell.Value;

            Cell = WSheet.Cells[Cell.Row + 1, Cell.Column];
            DataForm.Organization2.CompanyEmployee1.EmployeeName = Cell.Value;
            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
            DataForm.Organization2.CompanyEmployee1.EmployeePost = Cell.Value;
            Cell = WSheet.Cells[Cell.Row + 1, Cell.Column - 1];

            DataForm.Organization2.CompanyEmployee2.EmployeeName = Cell.Value;
            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
            DataForm.Organization2.CompanyEmployee2.EmployeePost = Cell.Value;
            Cell = WSheet.Cells[Cell.Row + 1, Cell.Column - 1];

            DataForm.Organization2.CompanyEmployee3.EmployeeName = Cell.Value;
            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
            DataForm.Organization2.CompanyEmployee3.EmployeePost = Cell.Value;

            CDH(12); // 3 Компания
            DataForm.Organization3.CompanyName = Cell.Value;

            Cell = WSheet.Cells[Cell.Row + 1, Cell.Column];
            DataForm.Organization3.CompanyEmployee1.EmployeeName = Cell.Value;
            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
            DataForm.Organization3.CompanyEmployee1.EmployeePost = Cell.Value;
            Cell = WSheet.Cells[Cell.Row + 1, Cell.Column - 1];

            DataForm.Organization3.CompanyEmployee2.EmployeeName = Cell.Value;
            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
            DataForm.Organization3.CompanyEmployee2.EmployeePost = Cell.Value;
            Cell = WSheet.Cells[Cell.Row + 1, Cell.Column - 1];

            DataForm.Organization3.CompanyEmployee3.EmployeeName = Cell.Value;
            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
            DataForm.Organization3.CompanyEmployee3.EmployeePost = Cell.Value;

            CDH(16);
            DataForm.BuildPlace1.HeightMark = Cell.Value;
            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
            DataForm.BuildPlace1.Axes = Cell.Value;

            CDH(17);
            DataForm.BuildPlace2.HeightMark = Cell.Value;
            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
            DataForm.BuildPlace2.Axes = Cell.Value;

            CDH(18);
            DataForm.BuildPlace3.HeightMark = Cell.Value;
            Cell = WSheet.Cells[Cell.Row, Cell.Column + 1];
            DataForm.BuildPlace3.Axes = Cell.Value;
        }
         
        //Заполнение полей документов
        public void PasteData(byte Template)
        {
            ///// Номер Документа
            Cell = WSheet.Cells.Find("НомерДокумента");
            Cell.Value = 1;

            ///// Объект Капитального Строительства
            Cell = WSheet.Cells.Find(((CellDocument)0).ToString());
            TextWrapping(DataForm[0], Cell);

            ///// ФИО подписантов
            for (byte i = 1; i < 10; i++)
            {
                Cell = WSheet.Cells.Find(((CellDocument)i).ToString());
                if (Cell != null)
                {
                    if (DataForm[i] != null)
                    {
                        Cell.Value = DataForm[i];
                        continue;
                    }
                    DeleteRows(Cell);
                }
            }

            ///// Подписанты
            for (byte i = 10; i < 19; i++)
            {
                Cell = WSheet.Cells.Find(((CellDocument)i).ToString());
                if (Cell != null)
                {
                    if (DataForm[i-9] != null)
                    {
                        Cell.Value = DataForm[i];
                        continue;
                    }
                    DeleteRows(Cell);
                }
            }

            //// Поля Каждого Шаблона
            switch (Template)
            {
                case 1:
                    FullReplace(((CellDocument)19).ToString(), DataForm[19]);
                    break; 
                case 2:
                    FullReplace(((CellDocument)21).ToString(), DataForm[21]);

                    Cell = WSheet.Cells.Find(((CellDocument)25).ToString());
                    string Value = "к производству работ по монтажу технических средств системы автоматической пожарной сигнализации, ";
                    for (byte i = 25; i < 31; i += 2)
                    {
                        if (DataForm[i] == null)
                            break;
                        Value += (DataForm[i] + " " + DataForm[i + 1] + ", ");
                    }
                    Value += ("согласно проекту " + DataForm[31]);
                    TextWrapping(Value, Cell, 112, 112);

                    break;
                case 3:
                    FullReplace(((CellDocument)22).ToString(), DataForm[22]);

                    Cell = WSheet.Cells.Find("МРВ");
                    if (Cell != null)
                    {
                        Cell.Replace("МРВ", DataForm.Organization3.CompanyName);
                    }

                    for (int i = 23; i < 25; i++)
                    {
                        Cell = WSheet.Cells.Find(((CellDocument)i).ToString());
                        if (Cell != null)
                        {
                            Cell.Replace(((CellDocument)i).ToString(), DataForm[i + 10]);
                        }
                    }

                    Cell = WSheet.Cells.Find(((CellDocument)26).ToString());
                    if (Cell != null)
                    {
                        TextWrapping($"1. {DataForm.Organization3.CompanyName} предъявлена к приемке система автоматической пожарной " +
                            $"сигнализации, смонтированная {DataForm.BuildingName}, {DataForm.BuildObject.Substring(DataForm.BuildObject.IndexOf("по адресу:"))}, " +
                            $"по проекту, разработанному {DataForm.ProjectCompany}, шифр проекта {DataForm.ProjectNumber}", Cell, 105, 110);
                    }

                    Cell = WSheet.Cells.Find(((CellDocument)27).ToString());
                    if (Cell != null)
                    {
                        TextWrapping($" Работы по монтажу выполнены в соответствии с проектом шифр {DataForm[31]}, разработанным {DataForm[32]}, стандартами, строительными нормами и правилами.", Cell, 100, 110);
                    }

                    Cell = WSheet.Cells.Find(((CellDocument)28).ToString());
                    if (Cell != null)
                    {
                        TextWrapping($" Системы пожарной безопасности, предъявленные к приемке, считать принятыми с {(DataForm.EndDate.AddDays(1).ToString("«dd» MMMM yyyy") + " г.")} для проведения пусконаладочных работ.", Cell, 100, 110);
                    }
                    break;
                case 4:
                    FullReplace(((CellDocument)22).ToString(), DataForm[22]);
                    break;
                case 5:
                    FullReplace("ДатаОкончания7", DataForm.EndDate.AddDays(7).ToString("«dd» MMMM yyyy") + " г.");

                    Cell = WSheet.Cells.Find("1Пункт");
                    if (Cell != null)
                    {
                        TextWrapping($"1. Монтажной организацией предъявлены к приемке система автоматической пожарной " +
                            $"сигнализации, смонтированная {DataForm.BuildingName}, {DataForm.BuildObject.Substring(DataForm.BuildObject.IndexOf("по адресу:"))}, " +
                            $"по проекту, разработанному {DataForm.ProjectCompany}, шифр проекта {DataForm.ProjectNumber}", Cell, 100, 100);
                    }

                    Cell = WSheet.Cells.Find("2Пункт");
                    if (Cell != null)
                    {
                        TextWrapping($"2. Монтажные работы выполнены {DataForm.Organization3.CompanyName} в период c {DataForm[33]} по {DataForm[34]}.", Cell, 97, 97);
                    }

                    Cell = WSheet.Cells.Find("ЗаключениеКомиссии");
                    if (Cell != null)
                    {
                        TextWrapping($"Систему автоматической пожарной сигнализации считать принятой в эксплуатацию с {DataForm.EndDate.AddDays(8).ToString("«dd» MMMM yyyy")} года.", Cell, 90, 90);
                    }
                    break;
                default:
                    break;
            }
        }
        public void PartReplace(string FindValue, string ReplaceValue)
        {
            Cell = WSheet.Cells.Find(FindValue);
            if (Cell != null)
            {
                Cell.Replace(FindValue, ReplaceValue);
            }
        }
        public void FullReplace(string FindValue, string ReplaceValue)
        {
            Cell = WSheet.Cells.Find(FindValue);
            if (Cell != null)
            {
                Cell.Value = ReplaceValue;
            }
        }
        #endregion

        #region Enumы Класса
        public enum CellDocument : byte
        {
            ОбъектКапитальногоСтроительства,
            ИИПредставительЗастройщикаN1,
            ИИПредставительЗастройщикаN2,
            ИИПредставительЗастройщикаN3,
            ИИПредставительГенеральногоПодрядчикаСтроительстваN1,
            ИИПредставительГенеральногоПодрядчикаСтроительстваN2,
            ИИПредставительГенеральногоПодрядчикаСтроительстваN3,
            ИИПредставительМонтажнойОрганизацииN1,
            ИИПредставительМонтажнойОрганизацииN2,
            ИИПредставительМонтажнойОрганизацииN3,
            ПредставительЗастройщикаN1,
            ПредставительЗастройщикаN2,
            ПредставительЗастройщикаN3,
            ПредставительГенеральногоПодрядчикаСтроительстваN1,
            ПредставительГенеральногоПодрядчикаСтроительстваN2,
            ПредставительГенеральногоПодрядчикаСтроительстваN3,
            ПредставительМонтажнойОрганизацииN1,
            ПредставительМонтажнойОрганизацииN2,
            ПредставительМонтажнойОрганизацииN3,
            ДатаНачалаРабот,
            ДатаОкончанияРабот,
            ДатаНачалаРаботПолнаяГ,
            ДатаОкончанияРаботПолнаяГ,
            ДатаНачалаРаботПолная,
            ДатаОкончанияРаботПолная,
            АГЗКПМР,
            АОМР,
            АОМРЗ1,
            АОМРЗ2,
            МРВ,
        }
        public enum CellData : byte
        {
            ОбъектКапитальногоСтроительства,
            ДатаНачалаРабот,
            ДатаОкончанияРабот,
            Шифрпроекта,

            Застройщик,
            ПредставительЗастройщикаN1,
            ПредставительЗастройщикаN2,
            ПредставительЗастройщикаN3,

            ГенеральныйПодрядчикСтроительства,
            ПредставительГенеральногоПодрядчикаСтроительстваN1,
            ПредставительГенеральногоПодрядчикаСтроительстваN2,
            ПредставительГенеральногоПодрядчикаСтроительстваN3,

            МонтажнаяОрганизация,
            ПредставительМонтажнойОрганизацииN1,
            ПредставительМонтажнойОрганизацииN2,
            ПредставительМонтажнойОрганизацииN3,

            План1,
            План2,
            План3,

            ПроектРазработал,
            НазваниеЗдания,
        }
        public enum Templates : byte
        {
            Акт_Входного_Контроля = 1,
            Акт_Готовности_Зданий_К_Производству_МР,
            Акт_Окончания_Монтажных_Работ,
            Ведомость_Смонтированного,
            Акт_Приемки_Системы_В_Эксплуатацию,
        }
        #endregion
    }
}
