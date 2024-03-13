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
            Worksheet employeesSheet = workbook.Sheets["Сотрудники"]; // Получение листа "Сотрудники" из книги Excel
            Worksheet departmentsSheet = workbook.Sheets["Отделы"]; // Получение листа "Отделы" из книги Excel
            Worksheet tasksSheet = workbook.Sheets["Задачи"]; // Получение листа "Задачи" из книги Excel

            ReadingData readingData = new ReadingData();

            List<Employee> employees = readingData.ReadEmployees(employeesSheet);
            List<Department> departments = readingData.ReadDepartments(departmentsSheet);
            Dictionary<string, int> taskCountByEmployee = readingData.CalculateTaskCountByEmployee(tasksSheet, employees);

            WordDocumentGenerator generator = new WordDocumentGenerator(wordFilePath);
            generator.GenerateReport(departments, employees, taskCountByEmployee);

            workbook.Close(false);
            excelApp.Quit();

            // Открываем документ Word и завершаем выполнение программы
            System.Diagnostics.Process.Start("WINWORD.EXE", $"{wordFilePath}Отчет.docx");
        }
    }
}
