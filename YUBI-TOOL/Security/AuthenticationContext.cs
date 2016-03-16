using System;

namespace YUBI_TOOL.Security
{
    public class AuthenticationContext
    {
        public AuthenticationContext()
        {
        }
        private object loginAgent;

        public object LoginAgent
        {
            get { return loginAgent; }
            set { loginAgent = value; }
        }
        private DateTime? loginTime;

        public DateTime? LoginTime
        {
            get { return loginTime; }
            set { loginTime = value; }
        }

    }
}
