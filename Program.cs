using System;

namespace excel
{
    class Program
    {
        static void Main(string[] args)
        {
            var excel = new StudentSpace.ExcelIO();
            //  excel.FillWithDummyData();
            //   excel.save("StudentData.xlsx");
            excel.load("StudentData.xlsx");
            excel.print();
            var guid = Guid.NewGuid();
        }

    }

}



