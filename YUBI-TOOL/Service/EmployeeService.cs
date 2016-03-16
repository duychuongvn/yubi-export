using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using YUBI_TOOL.Model;
using System.ComponentModel.Composition;
using System.Data.Linq.SqlClient;

namespace YUBI_TOOL.Service
{
    [Export(typeof(IEmployeeService))]
    public class EmployeeService : IEmployeeService
    {
        public EmployeeModel GetEmployee(decimal companyNo, decimal postNo, string employeeNo, decimal expirationFrom, decimal expirationTo)
        {
            EmployeeModel employeeModel = null;
            var context = Dao.DaoHelper.GetContext();
            var results = (from employee in context.EMPLOYEEs
                           join post in context.POSTs
                           on employee.POST_NO equals post.POST_NO
                           where employee.COMPANY_NO == companyNo
                           && post.COMPANY_NO == companyNo
                               // && employee.STATUS == Common.DBConstant.STATUS_ADD
                            && employee.EMPLOYEE_NO == employeeNo
                           && (employee.EXPIRATION_FROM <= expirationFrom && employee.EXPIRATION_TO >= expirationFrom
                           || employee.EXPIRATION_FROM <= expirationTo && employee.EXPIRATION_TO >= expirationTo)
                           orderby employee.EXPIRATION_TO descending
                           select new { employee, post });
                           
            
            var result = results.FirstOrDefault(); 
            if (result != null)
            {
                employeeModel = new EmployeeModel()
                {
                    Employee_no = result.employee.EMPLOYEE_NO,
                    Alphabet_first_name = result.employee.ALPHABET_FIRST_NAME,
                    Alphabet_last_name = result.employee.ALPHABET_LAST_NAME,
                    Company_no = result.employee.COMPANY_NO,
                    Create_date_time = result.employee.CREATE_DATE_TIME,
                    Emsize_first_name = result.employee.EMSIZE_FIRST_NAME,
                    Emsize_last_name = result.employee.EMSIZE_LAST_NAME,
                    Etirement_date = result.employee.RETIREMENT_DATE,
                    Expiration_from = result.employee.EXPIRATION_FROM,
                    Expiration_to = result.employee.EXPIRATION_TO,
                    Id = result.employee.ID,
                    Login_password = result.employee.LOGIN_PASSWORD.ToString(),
                    Post_no = result.employee.POST_NO,
                    Remarks = result.employee.REMARKS,
                    Status = result.employee.STATUS,
                    Time_table_no = result.employee.TIME_TABLE_NO,
                    Update_date_time = result.employee.UPDATE_DATE_TIME,
                    Use_flag_of_holiday = result.employee.USE_FLAG_OF_HOLIDAY,
                    Post_name = result.post.POST_NAME,
                };
            }
            return employeeModel;
        }
        public List<EmployeeModel> SearchEmployeeList(decimal companyNo, decimal postNo, string employeeName, decimal dateFrom, decimal dateTo)
        {
            List<EmployeeModel> employeeList = new List<EmployeeModel>();
            var context = Dao.DaoHelper.GetContext();
            var results = from employee in context.EMPLOYEEs
                          join post in context.POSTs
                          on employee.POST_NO equals post.POST_NO
                          where employee.COMPANY_NO == companyNo
                          && post.COMPANY_NO == companyNo
                          //&& employee.STATUS == Common.DBConstant.STATUS_ADD
                          //&& employee.POST_NO > 0
                          && (employee.EXPIRATION_FROM <= dateFrom && employee.EXPIRATION_TO >= dateFrom
                          || employee.EXPIRATION_FROM <= dateTo && employee.EXPIRATION_TO >= dateTo)
                          && (SqlMethods.Like(employee.EMPLOYEE_NO, string.Format("%{0}%", employeeName))
                           || SqlMethods.Like(employee.EMSIZE_FIRST_NAME + employee.EMSIZE_LAST_NAME, string.Format("%{0}%", employeeName))
                           ) 
                          orderby employee.POST_NO ascending, employee.EMPLOYEE_NO ascending
                          select new {employee, post};
            if (postNo > 0)
            {
                var results2 = results.Where(x => x.employee.POST_NO == postNo);
                foreach (var result in results2)
                {
                    EmployeeModel employeeModel = new EmployeeModel()
                    {
                        Company_no = result.employee.COMPANY_NO,
                        Alphabet_first_name = result.employee.EMSIZE_FIRST_NAME,
                        Alphabet_last_name = result.employee.EMSIZE_LAST_NAME,
                        Create_date_time = result.employee.CREATE_DATE_TIME,
                        Employee_no = result.employee.EMPLOYEE_NO,
                        Emsize_first_name = result.employee.EMSIZE_FIRST_NAME,
                        Emsize_last_name = result.employee.EMSIZE_LAST_NAME,
                        Etirement_date = result.employee.RETIREMENT_DATE,
                        Expiration_from = result.employee.EXPIRATION_FROM,
                        Expiration_to = result.employee.EXPIRATION_TO,
                        Id = result.employee.ID,
                        Login_password = result.employee.LOGIN_PASSWORD.ToString(),
                        Post_no = result.employee.POST_NO,
                        Remarks = result.employee.REMARKS,
                        Status = result.employee.STATUS,
                        Time_table_no = result.employee.TIME_TABLE_NO,
                        Update_date_time = result.employee.UPDATE_DATE_TIME,
                        Use_flag_of_holiday = result.employee.USE_FLAG_OF_HOLIDAY,
                        Post_name = result.post.POST_NAME
                    };
                    employeeList.Add(employeeModel);
                }
            }
            else
            {
                foreach (var result in results)
                {
                    EmployeeModel employeeModel = new EmployeeModel()
                    {
                        Company_no = result.employee.COMPANY_NO,
                        Alphabet_first_name = result.employee.EMSIZE_FIRST_NAME,
                        Alphabet_last_name = result.employee.EMSIZE_LAST_NAME,
                        Create_date_time = result.employee.CREATE_DATE_TIME,
                        Employee_no = result.employee.EMPLOYEE_NO,
                        Emsize_first_name = result.employee.EMSIZE_FIRST_NAME,
                        Emsize_last_name = result.employee.EMSIZE_LAST_NAME,
                        Etirement_date = result.employee.RETIREMENT_DATE,
                        Expiration_from = result.employee.EXPIRATION_FROM,
                        Expiration_to = result.employee.EXPIRATION_TO,
                        Id = result.employee.ID,
                        Login_password = result.employee.LOGIN_PASSWORD.ToString(),
                        Post_no = result.employee.POST_NO,
                        Remarks = result.employee.REMARKS,
                        Status = result.employee.STATUS,
                        Time_table_no = result.employee.TIME_TABLE_NO,
                        Update_date_time = result.employee.UPDATE_DATE_TIME,
                        Use_flag_of_holiday = result.employee.USE_FLAG_OF_HOLIDAY,
                        Post_name = result.post.POST_NAME

                    };
                    employeeList.Add(employeeModel);
                }
            }

            return employeeList;
        }
    }
}
