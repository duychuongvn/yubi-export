using System;
using System.ComponentModel.Composition;
using System.Linq;

namespace YUBI_TOOL.Security
{
    [Export(typeof(IAuthenticationManager))]
    public class AuthenticationManager : IAuthenticationManager
    {

        public AuthenticationManager()
        {
            AuthenticationContext = new AuthenticationContext();
        }
        public object DoAuthentication(string userName, string password)
        {
            var context = Dao.DaoHelper.GetContext();
            Model.EmployeeModel employeeModel = null;
            var employee = (from e in context.EMPLOYEEs
                            where e.EMPLOYEE_NO == userName
                             && e.STATUS == Common.DBConstant.STATUS_ADD
                            select e).FirstOrDefault();
            if (employee != null && employee.LOGIN_PASSWORD != null && employee.LOGIN_PASSWORD.ToString() == password)
            {

                employeeModel = new Model.EmployeeModel()
                {
                    Employee_no = employee.EMPLOYEE_NO,
                    Emsize_first_name = employee.EMSIZE_FIRST_NAME,
                    Emsize_last_name = employee.EMSIZE_LAST_NAME,
                    Login_password = employee.LOGIN_PASSWORD != null ? employee.LOGIN_PASSWORD.ToString() : null,
                };
                employeeModel.Employee_no = employee.EMPLOYEE_NO;

            }

            AuthenticationContext.LoginAgent = employeeModel;
            AuthenticationContext.LoginTime = DateTime.Now;
            return employeeModel;
        }

        public bool IsAuthenticated()
        {
            return AuthenticationContext.LoginAgent != null && AuthenticationContext.LoginTime != null;
        }
        public void DeAuthentication(string userName, string password)
        {
            AuthenticationContext.LoginAgent = null;
            AuthenticationContext.LoginTime = null;
        }
        public AuthenticationContext AuthenticationContext { get; set; }
    }
}
