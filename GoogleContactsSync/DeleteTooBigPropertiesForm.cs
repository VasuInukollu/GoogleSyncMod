using System.Drawing;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    public partial class DeleteTooBigPropertiesForm : Form
    {
        public DeleteTooBigPropertiesForm()
        {
            /* Cannot set Font in designer as there is automatic sorting and Font will be set after AutoScaleDimensions
             * This will prevent application to work correctly with high DPI systems. */
            Font = new Font("Verdana", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);

            InitializeComponent();       
        }
    }
}
