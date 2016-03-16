using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using YUBI_TOOL.Common;
using YUBI_TOOL.Model;

namespace YUBI_TOOL.Service
{
    [Export(typeof(IDepartmentService))]
    public class DepartmentService : IDepartmentService
    {
        public List<PostModel> SearchDepartment(decimal company_no)
        {
            List<PostModel> departmentList = new List<PostModel>();
            decimal currentDate = CommonUtil.GetCurrentDate();
            var context = Dao.DaoHelper.GetContext();
            var results = from post in context.POSTs
                          where post.COMPANY_NO == company_no && post.STATUS == DBConstant.STATUS_ADD
                          && post.EXPIRATION_FROM <= currentDate && post.EXPIRATION_TO >= currentDate
                          select post;
            foreach (var post in results)
            {
                PostModel postModel = new PostModel()
                {
                    Company_no = post.COMPANY_NO.ToString(),
                    Create_date_time = post.CREATE_DATE_TIME,
                    Expiration_from = post.EXPIRATION_FROM,
                    Expiration_to = post.EXPIRATION_TO,
                    Post_name = post.POST_NAME,
                    Post_no = post.POST_NO.ToString(),
                    Status = post.STATUS,
                    Update_date_time = post.UPDATE_DATE_TIME,
                   

                };
                departmentList.Add(postModel);
            }
            return departmentList;
        }
    }
}
