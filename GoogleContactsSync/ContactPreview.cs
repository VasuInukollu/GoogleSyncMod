using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.ObjectModel;

namespace GoContactSyncMod
{
    public partial class ContactPreview : UserControl
    {
        private Collection<CPField> fields;

        private Outlook.ContactItem outlookContact;
        public Outlook.ContactItem OutlookContact
        {
            get { return outlookContact; }
            set { outlookContact = value; }
        }

        public ContactPreview(Outlook.ContactItem _outlookContact)
        {
            InitializeComponent();
            outlookContact = _outlookContact;
            InitializeFields();
        }

        private void InitializeFields()
        {
            // TODO: init all non null fields
            fields = new Collection<CPField>();

            int index = 0;
            int height = Font.Height;

            if (outlookContact.FirstName != null)
            {
                fields.Add(new CPField("First name", outlookContact.FirstName, new PointF(0, index * height)));
                index++;
            }
            if (outlookContact.LastName != null)
            {
                fields.Add(new CPField("Last name", outlookContact.LastName, new PointF(0, index * height)));
                index++;
            }
            if (outlookContact.Email1Address != null)
            {
                fields.Add(new CPField("Email", ContactPropertiesUtils.GetOutlookEmailAddress1(outlookContact), new PointF(0, index * height)));
                index++;
            }
            
            // resize to fit
            this.Height = (index + 1) * height;
        }

        private void ContactPreview_Paint(object sender, PaintEventArgs e)
        {
            foreach (CPField field in fields)
                field.Draw(e, Font);
        }


    }

    public class CPField
    {
        private string Name;
        private string Value;
        private PointF P;

        public string name
        {
            get { return Name; }
            set { Name = value; }
        }

        public string value
        {
            get { return Value; }
            set { Value = value; }
        }

        public PointF p
        {
            get { return P; }
            set { P = value; }
        }


        public CPField(string nameVal, string valueVal, PointF pVal)
        {
            Name = nameVal;
            Value = valueVal;
            P = pVal;
        }

        public void Draw(PaintEventArgs e, Font font)
        {
            string str = Name + ": " + Value;
            if (e != null)
                e.Graphics.DrawString(str, font, Brushes.Black, P);
            else
                throw new ArgumentNullException("PaintEventArgs is null");
        }
    }
}
