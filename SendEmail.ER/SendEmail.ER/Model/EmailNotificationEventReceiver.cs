using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using SendEmail.ER.Controllers;

namespace SendEmail.ER.Model
{
    class EmailNotificationEventReceiver : SPItemEventReceiver
    {
        public override void ItemAdded(SPItemEventProperties properties)
        {
            EmailNotificationController.CheckSend(properties.SiteId, properties.Site.SystemAccount.UserToken, properties.Web.ID, properties.ListTitle, properties.ListItemId);
        }
    }
}
