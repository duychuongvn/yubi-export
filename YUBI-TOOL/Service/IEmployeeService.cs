using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using YUBI_TOOL.Model;

namespace YUBI_TOOL.Service
{
    public interface IEmployeeService
    {
        List<EmployeeModel> SearchEmployeeList(decimal companyNo, decimal postNo, string employeeName, decimal dateFrom, decimal dateTo);
        EmployeeModel GetEmployee(decimal companyNo, decimal postNo, string employeeNo, decimal expirationFrom, decimal expirationTo);
        
    }
}
