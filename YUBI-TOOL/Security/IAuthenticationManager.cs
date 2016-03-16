using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YUBI_TOOL.Security
{
    public interface IAuthenticationManager
    {
        object DoAuthentication(string userName, string password);
        void DeAuthentication(string userName, string password);
        bool IsAuthenticated();
    }
}
