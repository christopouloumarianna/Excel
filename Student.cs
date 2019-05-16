using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace StudentSpace
{
    public interface IExcel
    {
        bool save(string ExcelFileName);
        bool load(string ExcelFileName);
    }



    public class ExcelIO : IExcel
    {

        public List<Student> students = new List<Student>();
        public void FillWithDummyData()
        {
            students.Add(new Student("Dimitris", "BSc CS", 1));

            students.Add(new Student("Antonis", "MSc CS", 2));

            students.Add(new Student("Eugenia", "MSc ML", 3));
        }



        public bool save(string ExcelFileName)
        {

            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Mysheet");
            var r1 = sheet.CreateRow(0);
            r1.CreateCell(0).SetCellValue("Name");
            r1.CreateCell(1).SetCellValue("Course");
            r1.CreateCell(2).SetCellValue("RegistryId");
            for (int i = 0; i < students.Count; i++)
            {

                var r = sheet.CreateRow(i + 1);
                r.CreateCell(0).SetCellValue(students[i].Name);
                r.CreateCell(1).SetCellValue(students[i].Course);
                r.CreateCell(2).SetCellValue(students[i].RegistryId);



            }
            using (var fs = new FileStream(ExcelFileName, FileMode.Create,
            FileAccess.Write))
            {
                wb.Write(fs);
            }
            return true;
        }

        public bool load(string ExcelFileName)
        {
            students.Clear();
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(ExcelFileName, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
            ISheet sheet = hssfwb.GetSheet("Mysheet");
            //first line contains header
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null)
                {
                    string Name = sheet.GetRow(row).GetCell(0).ToString();
                    string Course = sheet.GetRow(row).GetCell(1).ToString();
                    int Id = int.Parse(sheet.GetRow(row).GetCell(2).ToString());
                    students.Add(new Student(Name, Course, Id));

                }

            }
            return true;

        }
        public void print()
        {
            foreach (var s in students)
            {
                Console.WriteLine(s);
            }
        }

    }
    public class Student
    {
        public int RegistryId { get; set; }
        public string Name { get; set; }
        public string Course { get; set; }

        public Student(string name, string course, int registryId)

        {

            Name = name;

            Course = course;

            RegistryId = registryId;

        }
    }
}


