using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using YUBI_TOOL.Model;

namespace YUBI_TOOL.Service
{
    public interface ITimeTableService
    {
        List<TimeTableModel> SearchTimeTableList();
    }
}
