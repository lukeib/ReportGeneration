using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ReportGeneration
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath;
            
            while (true)
            {
                Console.WriteLine("Введите абсолютный путь к Excel файлу " + "Data.xlsb"+":");
                excelFilePath = Console.ReadLine();
                if (File.Exists(excelFilePath))
                    break;
                Console.WriteLine("Ошибка: По указанному пути файл Data.xlsb не найден.");
            }
            Console.WriteLine("Введите путь для сохранения файла Word:");
            string wordFilePath = Console.ReadLine();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Worksheet employeesSheet = workbook.Sheets["Сотрудники"];
            Worksheet departmentsSheet = workbook.Sheets["Отделы"];
            Worksheet tasksSheet = workbook.Sheets["Задачи"];

            ReadingData dataProcessor = new ReadingData();

            List<Employee> employees = dataProcessor.ReadEmployees(employeesSheet);
            List<Department> departments = dataProcessor.ReadDepartments(departmentsSheet);
            Dictionary<string, int> taskCountByEmployee = dataProcessor.CalculateTaskCountByEmployee(tasksSheet, employees);

            WordDocumentGenerator generator = new WordDocumentGenerator(wordFilePath);
            generator.GenerateReport(departments, employees, taskCountByEmployee);

            workbook.Close(false);
            excelApp.Quit();
        }
    }
}
