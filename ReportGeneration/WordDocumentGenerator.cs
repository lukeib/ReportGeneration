using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportGeneration
{
    public class WordDocumentGenerator
    {
        private string _wordFilePath;
        public WordDocumentGenerator(string wordFilePath)
        {
            _wordFilePath = wordFilePath;
        }
        /// <summary>
        /// Генерация отчёта в формате word и создание таблицы.
        /// Результатом выполнения будет создан Word файл с таблицей состоящей из сотрудников, отделов и количества задач.
        /// </summary>
        /// <param name="departments">Список отделов</param>
        /// <param name="employees">Список сотрудников</param>
        /// <param name="taskCountByEmployee">Словарь с количеством задач</param>
        public void GenerateReport(List<Department> departments, List<Employee> employees, Dictionary<string, int> taskCountByEmployee)
        {
            departments = departments.OrderByDescending(department =>
                            employees.Where(emp => emp.DepartmentId == department.DepartmentId)
                            .Sum(emp => taskCountByEmployee.ContainsKey(emp.EmployeeId) ? taskCountByEmployee[emp.EmployeeId] : 0)).ToList();

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = wordApp.Documents.Add();
            Paragraph title = doc.Paragraphs.Add();
            title.Range.Text = "Отчет по загрузке";
            title.Range.Font.Name = "Calibri";
            title.Range.Font.Size = 14;
            title.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            title.Range.InsertParagraphAfter();
            Microsoft.Office.Interop.Word.Range paragraphRange = title.Range.Paragraphs.Last.Range;

            //Добавление двух абзачев для отсутпа заголовка от таблицы, если подумать то можно упростить.
            paragraphRange.InsertParagraphAfter();
            paragraphRange.InsertParagraphAfter();
            Table table = doc.Tables.Add(title.Range, departments.Count + employees.Count + 1 , 2);

            table.Cell(1, 1).Range.Text = "Отдел";
            table.Cell(1, 2).Range.Text = "Количество задач";

            for (int col = 1; col <= 2; col++)
            {
                var cell = table.Cell(1, col).Range;
                cell.Font.Bold = 1;
                cell.Font.Color = WdColor.wdColorWhite;
                cell.Shading.BackgroundPatternColor = WdColor.wdColorGray50;
            }

            int rowIndex = 2; 
            foreach (var department in departments)
            {
                //Подсчет общего количества задач для текущего отдела.
                int departmentTaskCount = employees
                    .Where(emp => emp.DepartmentId == department.DepartmentId)
                    .Sum(emp => taskCountByEmployee.ContainsKey(emp.EmployeeId) ? taskCountByEmployee[emp.EmployeeId] : 0);

                //Заполняем ячейку в таблице с названием отдела и количестов задач в отделе.
                table.Cell(rowIndex, 1).Range.Text = department.DepartmentName;
                table.Cell(rowIndex, 2).Range.Text = departmentTaskCount.ToString();
                table.Cell(rowIndex, 1).Range.Font.Bold = 1;
                table.Cell(rowIndex, 2).Range.Font.Bold = 1;
                table.Rows[rowIndex].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray15;

                rowIndex++; 

                foreach (var employee in employees.Where(emp => emp.DepartmentId == department.DepartmentId)
                                                   .OrderByDescending(emp => taskCountByEmployee.ContainsKey(emp.EmployeeId) ? taskCountByEmployee[emp.EmployeeId] : 0))
                {
                    //Заполняем ячейки таблицы данными о сотруднике и количестве его задач.
                    if (employee.Patronymic == null)
                    {
                        //Если отчество отсутствует то записываем только фамилию и инициал имени.
                        table.Cell(rowIndex, 1).Range.Text = $"{employee.LastName} {employee.FirstName.Substring(0, 1)}. ";
                        table.Cell(rowIndex, 2).Range.Text = taskCountByEmployee.ContainsKey(employee.EmployeeId) ? taskCountByEmployee[employee.EmployeeId].ToString() : "0";
                        rowIndex++; 
                    }
                    else
                    {
                        //Если отчество есть, то записываем полные инициалы.
                        table.Cell(rowIndex, 1).Range.Text = $"{employee.LastName} {employee.FirstName.Substring(0, 1)}. {employee.Patronymic.Substring(0, 1)}.";
                        table.Cell(rowIndex, 2).Range.Text = taskCountByEmployee.ContainsKey(employee.EmployeeId) ? taskCountByEmployee[employee.EmployeeId].ToString() : "0";
                        rowIndex++; 
                    }
                }
            }
            // Устанавливаем шрифт и размер текста для всех ячеек таблицы.
            for (int i = 1; i <= table.Rows.Count; i++)
            {
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    table.Cell(i, j).Range.Font.Name = "Calibri";
                    table.Cell(i, j).Range.Font.Size = 11;
                }
            }
            // Выравниваем текст в столбце "Отделы" по левому краю.
            for (int row = 2; row <= table.Rows.Count; row++)
            {
                table.Cell(row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            }
            for (int i = 1; i <= table.Rows.Count; i++)
            {
                var row = table.Rows[i];

                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    var cell = row.Cells[j];
                    string cellText = cell.Range.Text.Trim();

                    if (!string.IsNullOrEmpty(cellText))
                    {
                        cell.Borders.Enable = 1;
                    }
                }
            }
            // Сохраняем документ и отлавливаем возможные ошибки при сохрании (отсутствие доступа и т.д.).
            bool success  = false;
            while (!success)
            {
                try
                {
                    doc.SaveAs2($"{_wordFilePath}Отчет.docx");
                    wordApp.Quit();
                    Console.WriteLine($"Файл успешно сохранён по пути: {_wordFilePath}Отчет.docx");
                    success = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при сохранении файла: {ex.Message}");
                    Console.WriteLine("Для повторной попытки сохранения введите действие (Y/N):");
                    string response = Console.ReadLine();

                    if (response.ToUpper() != "Y")
                        break;
                }
            }
        }
    }
}
