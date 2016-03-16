using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using YUBI_TOOL.Model;
using YUBI_TOOL.Common;
using System.ComponentModel.Composition;

namespace YUBI_TOOL.Service
{
    [Export(typeof(ICompanyService))]
    public class CompanyService :ICompanyService
    {
        public List<CompanyModel> SearchCompanyList()
        {
            List<CompanyModel> companyList = new List<CompanyModel>();
            var context = Dao.DaoHelper.GetContext();
            var companyResult = from company in context.COMPANies
                                where company.STATUS == DBConstant.STATUS_ADD
                                select company;
            foreach (var company in companyResult)
            {
                CompanyModel companyModel = new CompanyModel
                {
                    Company_no = company.COMPANY_NO.ToString(),
                    Company_name = company.COMPANY_NAME,

                };
                companyList.Add(companyModel);
            }
            return companyList;
        }
    }
}
