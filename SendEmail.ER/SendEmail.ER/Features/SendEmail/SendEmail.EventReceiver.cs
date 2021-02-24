using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using SendEmail.ER.Utils;
using SendEmail.ER.Model;


namespace SendEmail.ER.Features.SendEmail
{
    /// <summary>
    /// This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
    /// </summary>
    /// <remarks>
    /// The GUID attached to this class may be used during packaging and should not be modified.
    /// </remarks>

    [Guid("056dba9c-725c-44ff-b64b-3009e913f6a2")]
    public class SendEmailEventReceiver : SPFeatureReceiver
    {
        // Uncomment the method below to handle the event raised after a feature has been activated.
        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            var web = (SPWeb)properties.Feature.Parent;
            string listTitle = "Lists/" + Constants.Lists.NEWS;
            var list = web.GetList(SPUrlUtility.CombineUrl(web.ServerRelativeUrl, listTitle));
            EventReceiverController.ProvisionListEventReceiver("EmailNotificationEventReceiver - ItemAdded", list, typeof(EmailNotificationEventReceiver), SPEventReceiverType.ItemAdded, SPEventReceiverSynchronization.Asynchronous, 20000);
        }
        // Uncomment the method below to handle the event raised before a feature is deactivated.
        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            var web = (SPWeb)properties.Feature.Parent;
            string listTitle = "Lists/" + Constants.Lists.NEWS;
            var list = web.GetList(SPUrlUtility.CombineUrl(web.ServerRelativeUrl, listTitle));
            EventReceiverController.DeleteListEventReceiver(list, typeof(EmailNotificationEventReceiver), SPEventReceiverType.ItemAdded);
        }


        // Uncomment the method below to handle the event raised after a feature has been installed.

        //public override void FeatureInstalled(SPFeatureReceiverProperties properties)
        //{
        //}


        // Uncomment the method below to handle the event raised before a feature is uninstalled.

        //public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        //{
        //}

        // Uncomment the method below to handle the event raised when a feature is upgrading.

        //public override void FeatureUpgrading(SPFeatureReceiverProperties properties, string upgradeActionName, System.Collections.Generic.IDictionary<string, string> parameters)
        //{
        //}
    }
}
