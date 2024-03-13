using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportGeneration
{
    public class ReadingData
    {
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
        private string GetStringValue(Microsoft.Office.Interop.Excel.Range cell)
        {
            if (cell != null && cell.Value2 != null)
            {
                return cell.Value2.ToString();
            }
            return null; 
        }
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
        private string GetStringValueDepartment(Microsoft.Office.Interop.Excel.Range cell)
        {
            if (cell != null && cell.Value2 != null)
            {
                return cell.Value2.ToString();
            }
            return null;
        }
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
