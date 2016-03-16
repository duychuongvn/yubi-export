
namespace YUBI_TOOL.Common
{
    public class DBConstant
    {
        public const int STATUS_DEL = 1;
        public const int STATUS_ADD = 0;
        public const int DEFAULT_FONT_SIZE = 13;
        public const int DEFAULT_FONT_SIZE_JAPANESE = 14;
        public const int DEFAULT_FONT_SIZE_ENGLISH = 13;
        public const int DEFAULT_FONT_SIZE_VIETNAMESE = 13;
        public const int FLAG_OF_HOLIDAY_USE = 1;
        public const int FLAG_OF_HOLIDAY_UNUSE = 0;
        
        public const int WORK_DAY_TYPE_NORMAL = 2;
        public const int WORK_DAY_TYPE_NORMAL_HOLIDAY = 5 ;
        public const int WORK_DAY_TYPE_NATIONAL_HOLIDAY = 6;
       
        /// <summary>
        /// Normal Duty=1
        /// </summary>
        public const int WORK_TYPE_NORMAL = 1;
        /// <summary>
        /// Annual Leave=1
        /// </summary>
        public const int WORK_TYPE_ANNUAL_LEAVE = 2;
        /// <summary>
        /// Permission=7
        /// </summary>
        public const int WORK_TYPE_PERMISSION = 7;
        /// <summary>
        /// Absent without permission = 28
        /// </summary>
        public const int WORK_TYPE_AW_PERMISSION = 28;
        /// <summary>
        /// 1/2 Permission = 30
        /// </summary>
        public const int WORK_TYPE_HALF_PERMISSION = 30;

        /// <summary>
        /// Maternity Leave = 31
        /// </summary>
        public const int WORK_TYPE_MATERNITY_LEAVE = 31;
        /// <summary>
        /// Special Leave = 5
        /// </summary>
        public const int WORK_TYPE_SPECIAL_LEAVE = 5;
        /// <summary>
        /// Holiday Duty = 8
        /// </summary>
        public const int WORK_TYPE_HOLIDAY_DUTY = 8;

        public const int EMPL_REMARKS_LEN = 100;
        public const int EMPL_POSITION_LEN = 50;
        public const int EMPL_TEAM_LEN = 50;

        public const string PRESENCE_PRESENT = "o";
        public const string ABSENT_ANNUAL_LEAVE = "AL";
        public const string ABSENT_SPECIAL_LEAVE = "SL";
        public const string ABSENT_MATERNITY_LEAVE = "ML";
        public const string ABSENT_PERMISSION = "P";
        public const string ABSENT_HALF_PERMISSION = "1/2P";
        public const string ABSENT_WITHOUT_PERMISSION = "WP";
        public const string ABSENT_EXCLUDED = "-";
        public const string ABSENT_OFF = "OFF";
        public const string COMPANY_NO_ALL = "0";
        public const string COMPANY_NO_SPINNING_MILL = "1";
        public const string COMPANY_NO_KNITTING = "2";
        public const string POST_ALL_NO = "0";
        public const string CHAR_UNKNOWN = "-";
        public const int POST_SM_REGULAR_SHIFT = 1;
        public const int POST_SM_TEAM_A = 2;
        public const int POST_SM_TEAM_B = 3;
        public const int POST_SM_TEAM_C = 4;
        public const int POST_SM_TEAM_D = 5;
        public const int POST_SM_EXPEDITION = 6;
        public const int POST_SM_MECHANICS = 7;
        public const int POST_SM_ELECTRICIANS = 8;
        public const int POST_SM_REGULAR_SHIFT_ML = 9;

        public const int POST_KM_REGULAR_SHIFT = 2;
        public const int POST_KM_TEAM_A = 7;
        public const int POST_KM_TEAM_B = 8;
        public const int POST_KM_TEAM_C = 9;
        public const int POST_KM_TEAM_D = 10;
        public const int POST_KM_TEAM_A_ML = 28;
        public const int POST_KM_TEAM_B_ML = 29;

        public const int TIME_TABLE_NO_DEFAULT = 1;
        public const string COLOR_GRAY = "#969696";
        public const string COLOR_WHITE = "#FFFFFF";
        public const string COLOR_RED = "Red";
        public const string COLOR_HOLIDAY_NATIONAL = "#99CC00";
        public const string COLOR_HOLIDAY_WEEKEND = "#FF99CC";
        public const string COLOR_HOLIDAY_XLS = "#F2DDDC";
        public const int HOLIDAY_FLAG_NATIONAL = 1;
        public const int HOLIDAY_FLAG_WEEKEND = 0;

        public const string TEMP_DAILY_REPORT_BY_DEPART = "Dailly report each department";
        public const string TEMP_DAILY_REPORT_BY_COMPANY = "Dailly report all company";
        public const string SHEET_DAILY_REPORT = "Daily Report";
        public const string TEMP_DAILY_REPORT_NO_ON_DUTY = "No on duty";
        public const string TEMP_DAILY_REPORT_NO_OFF_DUTY = "No off duty";
        public const string TEMP_DAILY_REPORT_LATE = "Late";
        public const string TEMP_DAILY_REPORT_LEFT_EARLIER = "Left earlier";
        public const string TEMP_DAILY_RECAP_1 = "RECAP 1";
        public const string TEMP_DAILY_RECAP_2 = "RECAP 2";
        public const string TEMP_MONTHLY_ATTENDANCE_BY_DEPART = "Attendance for each deparment";
        public const string TEMP_MONTHLY_ATTENDANCE_BY_COMPANY = "Attendance for all company";
        public const string TEMP_MONTHLY_VIOLATION_BY_DEPART = "Violation for each department";
        public const string TEMP_MONTHLY_VIOLATION_BY_COMPANY = "Violation for all company";

        public const string SHEET_MONTHLY_VIOLATION = "Violation Monthly";
        public const string SHEET_MONTHLY_ATTENDANCE = "Attendance";

        public const string SHEET_RECAP_SPINNING_1 = "RECAP SM-1";
        public const string SHEET_RECAP_SPINNING_2 = "RECAP SM-2";
        public const string SHEET_RECAP_KNITTING_1 = "RECAP KM";
        public const string SHEET_RECAP_KNITTING_2 = "RECAP Knitting";
        public const string SHEET_RECAP_KNITTING_2_HEADER = "RECAPITULATION KNITTING";
        public const string SHEET_RECAP_SPINNING_2_HEADER = "RECAPITULATION SPINNING";
        public const string SHEET_RECAP_CONFECTION = "RECAP Confection";
        public const string SHEET_RECAP_EXPEDITION = "RECAP Expedition";
        public const string SHEET_DEFAULT = "Sheet1";
        public const int MAX_SHEET_PER_WORKSPACE = 250;
        public const int MAX_EMPLOYEE_NO_PER_FILE_NAME = 10;
        public const string RECAP_2_NOT_SAME_SHIFT_DISP = "Day off";
        public const string RECAP_2_START_WITH_KNITTING_DEPRT = "KNITTING";
        public const string CONFECTION_START_WITH_KNITTING_DEPRT = "CONFECTION";
        public const string EXPEDITION_START_WITH_KNITTING_DEPRT = "EXPEDITION";
        public const string RECAP_2_DEPART_TEAM = "TEAM";
        public const decimal WORKING_HALF_PERMISION = 0.5M;
        public const int PROGRAM_START_YEAR = 2012;
    }
}
