using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.ObjectModel;
using System.Drawing;

namespace GoContactSyncMod
{
    public partial class DeleteDuplicatesForm : Form
    {
        private Collection<ContactPreview> previews;

        public DeleteDuplicatesForm(Collection<Outlook.ContactItem> outlookContacts)
        {
            /* Cannot set Font in designer as there is automatic sorting and Font will be set after AutoScaleDimensions
             * This will prevent application to work correctly with high DPI systems. */
            Font = new Font("Verdana", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);

            InitializeComponent();
            previews = new Collection<ContactPreview>();
            Collection<Outlook.ContactItem> duplicates = FindDuplicates(outlookContacts);

            foreach (Outlook.ContactItem outlookContact in duplicates)
            {
                using (ContactPreview preview = new ContactPreview(outlookContact))
                {
                    preview.Parent = flowLayoutPanel;
                    previews.Add(preview);
                }
            }
        }

        private static Collection<Outlook.ContactItem> FindDuplicates(Collection<Outlook.ContactItem> outlookContacts)
        {
            // TODO:
            return null;
        }
    }
}