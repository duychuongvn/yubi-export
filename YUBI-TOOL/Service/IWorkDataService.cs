using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using YUBI_TOOL.Model;

namespace YUBI_TOOL.Service
{
    public interface IWorkDataService
    {
        List<WorkDataModel> SearchWorkDataListNightShift(decimal companyNo, decimal dateForm, decimal dateTo);
        List<WorkDataModel> SearchWorkDataList(decimal companyNo, decimal postNo, string employee, decimal dateForm, decimal dateTo);
        List<WorkDataModel> SearchWorkDataListByEmployee(decimal companyNo, decimal postNo, string employeeNo, decimal dateForm, decimal dateTo);
        WorkDataModel SearchWorkData(decimal companyNo, string employeeNo, decimal workDate);
        void UpdateWorkDataList(List<WorkDataModel> workDataList);
        void CreateWorkDataInMonth(EmployeeModel employee, decimal yearMonthDate);
    }
}
