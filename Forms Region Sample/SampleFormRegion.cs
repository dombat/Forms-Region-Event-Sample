using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Forms_Region_Sample
{
    partial class SampleFormRegion
    {
        #region Form Region Factory 

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Appointment)]
        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("Forms Region Sample.FormRegion1")]
        public partial class FormRegion1Factory
        {
            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void FormRegion1Factory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
            }
        }
        #endregion

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void FormRegion1_FormRegionShowing(object sender, System.EventArgs e)
        {
            try
            {
                ItemProperty property = ((dynamic)OutlookItem).ItemProperties["THIS WILL CAUSE AN EXEPTION!!!"];
                //if var is used instead of ItemProperty it has different behaviour
            }
            catch (System.Exception)
            {
                //ignore - just to slow down the processing of this method
            }

            if (OutlookItem is MailItem) 
            {
                ((ItemEvents_10_Event)OutlookItem).Send += MarkingFormRegion_Send;
            }
            else
            {
                MessageBox.Show("Its not a mail item");
            }

        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void FormRegion1_FormRegionClosed(object sender, System.EventArgs e)
        {
            if (OutlookItem is MailItem) ((ItemEvents_10_Event)OutlookItem).Send -= MarkingFormRegion_Send;
        }

        private void MarkingFormRegion_Send(ref bool cancel)
        {
            MessageBox.Show("Sending");
        }
    }
}
