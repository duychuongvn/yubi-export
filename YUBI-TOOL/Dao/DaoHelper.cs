using System;

namespace YUBI_TOOL.Dao
{
    public class DaoHelper
    {
        public static YubitaroDataContext GetContext()
        {
            YubitaroDataContext context = new YubitaroDataContext();
            return context;
        }
    }
}
