using System;

namespace DocumentGenerator
{
    // Класс с подклассами для извлечения и хранения данных
    public class Data_Form
    {
        public Data_Form()
        {
            Organization1 = new Company();
            Organization2 = new Company();
            Organization3 = new Company();
            BuildPlace1 = new Building();
            BuildPlace2 = new Building();
            BuildPlace3 = new Building();
        }
        public string BuildObject { get; set; }
        public string BuildingName { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string ProjectNumber { get; set; }
        public string ProjectCompany { get; set; }
        public Company Organization1 { get; set; }
        public Company Organization2 { get; set; }
        public Company Organization3 { get; set; }
        public Building BuildPlace1 { get; set; }
        public Building BuildPlace2 { get; set; }
        public Building BuildPlace3 { get; set; }
        public string this[int index]
        {
            get
            {
                switch (index)
                {
                    case 0: return BuildObject;

                    case 1: return Organization1.CompanyEmployee1.EmployeeName;
                    case 2: return Organization1.CompanyEmployee2.EmployeeName;
                    case 3: return Organization1.CompanyEmployee3.EmployeeName;

                    case 4: return Organization2.CompanyEmployee1.EmployeeName;
                    case 5: return Organization2.CompanyEmployee2.EmployeeName;
                    case 6: return Organization2.CompanyEmployee3.EmployeeName;

                    case 7: return Organization3.CompanyEmployee1.EmployeeName;
                    case 8: return Organization3.CompanyEmployee2.EmployeeName;
                    case 9: return Organization3.CompanyEmployee3.EmployeeName;

                    case 10: return (Organization1.CompanyEmployee1.EmployeePost + " " + Organization1.CompanyName + " " + Organization1.CompanyEmployee1.EmployeeName);
                    case 11: return (Organization1.CompanyEmployee2.EmployeePost + " " + Organization1.CompanyName + " " + Organization1.CompanyEmployee2.EmployeeName);
                    case 12: return (Organization1.CompanyEmployee3.EmployeePost + " " + Organization1.CompanyName + " " + Organization1.CompanyEmployee3.EmployeeName);

                    case 13: return (Organization2.CompanyEmployee1.EmployeePost + " " + Organization2.CompanyName + " " + Organization2.CompanyEmployee1.EmployeeName);
                    case 14: return (Organization2.CompanyEmployee2.EmployeePost + " " + Organization2.CompanyName + " " + Organization2.CompanyEmployee2.EmployeeName);
                    case 15: return (Organization2.CompanyEmployee3.EmployeePost + " " + Organization2.CompanyName + " " + Organization2.CompanyEmployee3.EmployeeName);

                    case 16: return (Organization3.CompanyEmployee1.EmployeePost + " " + Organization3.CompanyName + " " + Organization3.CompanyEmployee1.EmployeeName);
                    case 17: return (Organization3.CompanyEmployee2.EmployeePost + " " + Organization3.CompanyName + " " + Organization3.CompanyEmployee2.EmployeeName);
                    case 18: return (Organization3.CompanyEmployee3.EmployeePost + " " + Organization3.CompanyName + " " + Organization3.CompanyEmployee3.EmployeeName);

                    case 19: return StartDate.ToShortDateString();
                    case 20: return EndDate.ToShortDateString();

                    case 21: return StartDate.ToString("«dd» MMMM yyyy") + " г.";
                    case 22: return EndDate.ToString("«dd» MMMM yyyy") + " г.";

                    case 33: return StartDate.ToString("«dd» MMMM yyyy") + " года";
                    case 34: return EndDate.ToString("«dd» MMMM yyyy") + " года";

                    case 25: return BuildPlace1.HeightMark;
                    case 26: return BuildPlace1.Axes;
                    case 27: return BuildPlace2.HeightMark;
                    case 28: return BuildPlace2.Axes;
                    case 29: return BuildPlace3.HeightMark;
                    case 30: return BuildPlace3.Axes;

                    case 31: return ProjectNumber;
                    case 32: return ProjectCompany;

                    default: return null;
                }
            }
        }
    }
    public class Company
    {
        public Company()
        {
            CompanyEmployee1 = new Employee();
            CompanyEmployee2 = new Employee();
            CompanyEmployee3 = new Employee();
        }

        public string CompanyName { get; set; }
        public Employee CompanyEmployee1 { get; set; }
        public Employee CompanyEmployee2 { get; set; }
        public Employee CompanyEmployee3 { get; set; }
    }
    public class Employee
    {
        public string EmployeeName { get; set; }
        public string EmployeePost { get; set; }
    }
    public class Building
    {
        private string axes;
        private string heightmark;
        public string Axes
        {
            get { return "в осях " + axes; }
            set { axes = value; }
        }

        public string HeightMark
        {
            get { return "на отм. " + heightmark; }
            set { heightmark = value; }
        }
    }
}
