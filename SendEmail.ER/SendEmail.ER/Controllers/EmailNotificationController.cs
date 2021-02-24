using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Specialized;
using Microsoft.SharePoint;
using System.Net.Mail;
using Microsoft.SharePoint.Administration;

namespace SendEmail.ER.Controllers
{
    class EmailNotificationController
    {
        internal static void CheckSend(Guid SiteId, SPUserToken UserToken, Guid WebId, string ListTitle, int ListItemId)
        {
            using (SPSite s = new SPSite(SiteId, UserToken))
            {
                using (SPWeb w = s.OpenWeb(WebId))
                {
                    try
                    {
                        SPList newList = w.Lists.TryGetList(ListTitle);
                        SPListItem currentItem = newList.GetItemById(ListItemId);
                        //Проверяем помечено ли для рассылки 
                        if ((bool)currentItem[Constants.NewsFileds.WITH_EMAIL] == true)
                        {
                            BeginSend(w, currentItem);
                        }
                        else
                        {
                            return;
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
        }
        private static void BeginSend(SPWeb w, SPListItem currentItem)
        {
            SPListItemCollection EmployeeList = GetEmployeeList(w, currentItem);
            if (EmployeeList.Count > 0)
            {
                int userCount = 0;
                foreach(SPListItem userItem in EmployeeList)
                {
                    bool isSend = TrySendMail(userItem, w, currentItem);
                    if (isSend)
                    {
                        userCount++;
                    }
                }
                SaveHistory(w, currentItem, userCount);
            }
        }
        private static void SaveHistory(SPWeb w, SPListItem currentItem, int userCount)
        {
            SPList historyList = w.Lists.TryGetList(Constants.Lists.SP_HISTORY);
            var addedItem = historyList.AddItem();
            addedItem[Constants.HistoryFileds.TITLE] = currentItem[Constants.NewsFileds.TITLE];
            addedItem[Constants.HistoryFileds.SEND_DATE] = Microsoft.SharePoint.Utilities.SPUtility.CreateISO8601DateTimeFromSystemDateTime(DateTime.Now);
            addedItem[Constants.HistoryFileds.COUNT_USERS] = userCount;
            addedItem.Update();
        }
        private static bool TrySendMail(SPListItem userItem, SPWeb w, SPListItem currentItem)
        {
            try
            {
                if (!string.IsNullOrEmpty(userItem[Constants.EmplInGroupFileds.EMAIL].ToString()))
                {
                    SPFieldLookupValue EmailLookup = new SPFieldLookupValue(userItem[Constants.EmplInGroupFileds.EMAIL].ToString());
                    string Email = EmailLookup.LookupValue;
                    var email_feedback = "news@portal.ru";
                    string Subject = "Новая новость";
                    var header = new StringDictionary
                    {
                        {"from", email_feedback},
                        {"to", Email},
                        {"subject", Subject},
                        {"content-type", "text/html"},
                        {"fHtmlEncode", "False"},
                        {"fAppendHtmlTag", "False"}
                    };
                    SmtpClient client = new SmtpClient
                    {

                        Port = 25,
                        Host = SPAdministrationWebApplication.Local.OutboundMailServiceInstance.Server.Address,
                        DeliveryMethod = SmtpDeliveryMethod.Network
                    };
                    MailMessage email = new MailMessage(header["from"], header["to"]);
                    email.Subject = header["subject"];
                    email.From = new MailAddress(header["from"], "Сайт новостей");
                    email.IsBodyHtml = true;
                    email.Body = currentItem[Constants.NewsFileds.BODY].ToString();
                    client.Send(email);
                    return true;
                }
                else
                {
                    return false;
                }
            }catch(Exception ex)
            {
                return false;
            }
        }

        private static SPListItemCollection GetEmployeeList(SPWeb w, SPListItem currentItem)
        {
            SPList EmplInGroup = w.Lists.TryGetList(Constants.Lists.EMPL_IN_GROUP);
            SPQuery q = new SPQuery();
            q.Query =
            @"<Where>
                <Eq>
                    <FieldRef Name='" + Constants.EmplInGroupFileds.GROUP_ID + @"' />
                    <Value Type='Lookup'>" + Constants.GroupType.USERS + @"</Value>
                </Eq>
             </Where>";
            q.ViewFields = @"<FieldRef Name='" + Constants.EmplInGroupFileds.EMAIL + @"' />";
            SPListItemCollection listItems = EmplInGroup.GetItems(q);

            return listItems;
        }
    }
}
