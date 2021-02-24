using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendEmail.ER
{
    class Constants
    {
        public static class NewsFileds
        {
            public const string TITLE = "Title";
            public const string BODY = "Body";
            public const string WITH_EMAIL = "WithEmail";
        }
        public static class HistoryFileds
        {
            public const string TITLE = "Title";
            public const string SEND_DATE = "SendDate";
            public const string COUNT_USERS = "CountUsers";
        }
        public static class Lists
        {
            public const string NEWS = "News";
            public const string SP_HISTORY = "SP History";
            public const string EMPL_IN_GROUP = "Empl_in_Group";
            public const string GROUP = "Group";
            public const string EMPLOYEE = "Employee";
        }
        public static class EmplInGroupFileds
        {
            public const string EMPLOYEE_ID = "EmployeeId";
            public const string GROUP_ID = "GroupId";
            //При создании lookup поля EMPLOYEE_ID, можно указать "Добавьте столбец для отображения каждого из этих дополнительных полей:"
            //с ссылкой на поле Email из списка Employee
            public const string EMAIL = "EmployeeId_x003a_Email";
        }
        public static class GroupType
        {
            public const string MANAGERS = "Руководство";
            public const string USERS = "Пользователи сайта";
        }
    }
}
