using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportGeneration
{
    public class ReadingData
    {
        /// <summary>
        /// Считываем данные о сотрудниках из листа "Сотрудники"
        /// </summary>
        /// <param name="employeesSheet">Лист "Сотрудники"</param>
        /// <returns>Список объектов типа Employee</returns>
        public List<Employee> ReadEmployees(Worksheet employeesSheet)
        {
            List<Employee> employees = new List<Employee>();

            Microsoft.Office.Interop.Excel.Range usedRange = employeesSheet.UsedRange;

            var query = from Microsoft.Office.Interop.Excel.Range row in usedRange.Rows
                        select new Employee
                        {
                            EmployeeId = GetStringValue(row.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range),
                            LastName = GetStringValue(row.Cells[1, 2] as Microsoft.Office.Interop.Excel.Range)?.Trim(),
                            FirstName = GetStringValue(row.Cells[1, 3] as Microsoft.Office.Interop.Excel.Range)?.Trim(),
                            Patronymic = GetStringValue(row.Cells[1, 4] as Microsoft.Office.Interop.Excel.Range)?.Trim(),
                            BirthDate = GetDateStringEmployee(row.Cells[1, 5] as Microsoft.Office.Interop.Excel.Range),
                            DepartmentId = GetIntValueEmployee(row.Cells[1, 6] as Microsoft.Office.Interop.Excel.Range)
                        };

            employees = query.Skip(1)
                            .Where(emp => !string.IsNullOrEmpty(emp.EmployeeId))
                            .ToList();

            return employees;
        }
        /// <summary>
        /// Получаем строку из ячейки Excel
        /// </summary>
        /// <param name="cell">Ячейка Excel</param>
        /// <returns>Строка или null, если ячейка пустая</returns>
        private string GetStringValue(Microsoft.Office.Interop.Excel.Range cell)
        {
            if (cell != null && cell.Value2 != null)
            {
                return cell.Value2.ToString();
            }
            return null; 
        }
        /// <summary>
        /// Получаем дату в формате короткой строки из ячейки Excel
        /// </summary>
        /// <param name="cell">Ячейка Excel</param>
        /// <returns>Дата в формате короткой сткроки или null, если ячейка не является датой</returns>
        private string GetDateStringEmployee(Microsoft.Office.Interop.Excel.Range cell)
        {
            if (cell != null && cell.Value2 != null)
            {
                double dateValue;
                if (double.TryParse(cell.Value2.ToString(), out dateValue))
                {
                    DateTime date = DateTime.FromOADate(dateValue);
                    return date.ToShortDateString();
                }
            }
            return null; 
        }
        /// <summary>
        /// Получаем целое число из ячейки
        /// </summary>
        /// <param name="cell">Ячейка Учсуд</param>
        /// <returns>Целове число или 0, если значение в ячейке не является целым числом</returns>
        private int GetIntValueEmployee(Microsoft.Office.Interop.Excel.Range cell)
        {
            if (cell != null && cell.Value2 != null)
            {
                double doubleValue;
                if (double.TryParse(cell.Value2.ToString(), out doubleValue))
                {
                    return (int)doubleValue;
                }
            }
            return 0; 
        }
        /// <summary>
        /// Считываем данные из листа "Отделы"
        /// </summary>
        /// <param name="departmentsSheet">Лист "Отделы"</param>
        /// <returns>Список объектов типа Department </returns>
        public List<Department> ReadDepartments(Worksheet departmentsSheet)
        {
            List<Department> departments = new List<Department>();

            Microsoft.Office.Interop.Excel.Range usedRange = departmentsSheet.UsedRange;

            var query = from Microsoft.Office.Interop.Excel.Range row in usedRange.Rows
                        where row.Row > 1 
                        select new Department
                        {
                            DepartmentId = GetIntValueDepartment(row.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range),
                            DepartmentName = GetStringValueDepartment(row.Cells[1, 2] as Microsoft.Office.Interop.Excel.Range)
                        };

            departments = query.ToList();
            return departments;
        }
        /// <summary>
        /// Получаем строку из ячейки
        /// </summary>
        /// <param name="cell">Ячейка Excel из которой извлекаем данные</param>
        /// <returns>Строка или null, если ячейка пустая</returns>
        private string GetStringValueDepartment(Microsoft.Office.Interop.Excel.Range cell)
        {
            if (cell != null && cell.Value2 != null)
            {
                return cell.Value2.ToString();
            }
            return null;
        }
        /// <summary>
        /// Извлечение целого числа из ячейки
        /// </summary>
        /// <param name="cell">Ячейка Excel</param>
        /// <returns>Целое число или 0, если ячейка пустая или не содержит челое число</returns>
        private int GetIntValueDepartment(Microsoft.Office.Interop.Excel.Range cell)
        {
            if (cell != null && cell.Value2 != null)
            {
                double doubleValue;
                if (double.TryParse(cell.Value2.ToString(), out doubleValue))
                {
                    return (int)doubleValue;
                }
            }
            return 0;
        }
        /// <summary>
        /// Подсчёт количества задач для каждого сторудника
        /// </summary>
        /// <param name="tasksSheet">Лист "Задчи"</param>
        /// <param name="employees">Список сотруников, для которых нужно посчитать кол-во задач</param>
        /// <returns>Словарь, где ключ - ядентификатор сотрудника, значние - количество задач</returns>
        public Dictionary<string, int> CalculateTaskCountByEmployee(Worksheet tasksSheet, List<Employee> employees)
        {
            Dictionary<string, int> taskCountByEmployee = new Dictionary<string, int>();

            Microsoft.Office.Interop.Excel.Range usedRange = tasksSheet.UsedRange;

            var query = from Microsoft.Office.Interop.Excel.Range row in usedRange.Rows
                        where row.Row > 1 
                        let employeeId = GetStringCellValueTask(row.Cells[1, 2] as Microsoft.Office.Interop.Excel.Range) 
                        where employees.Any(emp => emp.EmployeeId == employeeId) 
                        select employeeId; 

            foreach (var employeeId in query)
            {
                if (taskCountByEmployee.ContainsKey(employeeId))
                    taskCountByEmployee[employeeId]++;
                else
                    taskCountByEmployee[employeeId] = 1;
            }
            return taskCountByEmployee;
        }
        /// <summary>
        /// Получаем строку из ячейки
        /// </summary>
        /// <param name="cell">Ячейка Excel из которой производится считывание</param>
        /// <returns>Строка или null, если ячейка пустая</returns>
        private string GetStringCellValueTask(Microsoft.Office.Interop.Excel.Range cell)
        {
            if (cell != null && cell.Value2 != null)
            {
                return cell.Value2.ToString();
            }
            return null;
        }
    }
}
