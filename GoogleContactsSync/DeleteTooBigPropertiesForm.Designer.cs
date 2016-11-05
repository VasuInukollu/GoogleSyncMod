using System.ComponentModel;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    partial class DeleteTooBigPropertiesForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DeleteTooBigPropertiesForm));
            this.bbOK = new System.Windows.Forms.Button();
            this.bbCancel = new System.Windows.Forms.Button();
            this.propertiesGrid = new System.Windows.Forms.DataGridView();
            this.Selected = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Key = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Value = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.explanationLabel = new System.Windows.Forms.Label();
            this.allCheck = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.propertiesGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // bbOK
            // 
            this.bbOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.bbOK.AutoEllipsis = true;
            this.bbOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.bbOK.Location = new System.Drawing.Point(213, 407);
            this.bbOK.Name = "bbOK";
            this.bbOK.Size = new System.Drawing.Size(87, 23);
            this.bbOK.TabIndex = 1;
            this.bbOK.Text = "OK";
            this.bbOK.UseVisualStyleBackColor = false;
            // 
            // bbCancel
            // 
            this.bbCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.bbCancel.AutoEllipsis = true;
            this.bbCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.bbCancel.Location = new System.Drawing.Point(340, 407);
            this.bbCancel.Name = "bbCancel";
            this.bbCancel.Size = new System.Drawing.Size(87, 23);
            this.bbCancel.TabIndex = 2;
            this.bbCancel.Text = "Cancel";
            this.bbCancel.UseVisualStyleBackColor = true;
            // 
            // propertiesGrid
            // 
            this.propertiesGrid.AllowUserToAddRows = false;
            this.propertiesGrid.AllowUserToDeleteRows = false;
            this.propertiesGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.propertiesGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.propertiesGrid.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.propertiesGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.propertiesGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Selected,
            this.Key,
            this.Value});
            this.propertiesGrid.Location = new System.Drawing.Point(10, 74);
            this.propertiesGrid.Name = "propertiesGrid";
            this.propertiesGrid.RowHeadersVisible = false;
            this.propertiesGrid.Size = new System.Drawing.Size(610, 309);
            this.propertiesGrid.TabIndex = 3;
            // 
            // Selected
            // 
            this.Selected.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Selected.HeaderText = "";
            this.Selected.Name = "Selected";
            this.Selected.Width = 5;
            // 
            // Key
            // 
            this.Key.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Key.HeaderText = "Key";
            this.Key.Name = "Key";
            this.Key.ReadOnly = true;
            this.Key.Width = 54;
            // 
            // Value
            // 
            this.Value.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Value.HeaderText = "Value";
            this.Value.Name = "Value";
            this.Value.ReadOnly = true;
            this.Value.Width = 63;
            // 
            // explanationLabel
            // 
            this.explanationLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.explanationLabel.AutoEllipsis = true;
            this.explanationLabel.AutoSize = true;
            this.explanationLabel.Location = new System.Drawing.Point(10, 9);
            this.explanationLabel.MinimumSize = new System.Drawing.Size(606, 52);
            this.explanationLabel.Name = "explanationLabel";
            this.explanationLabel.Size = new System.Drawing.Size(606, 52);
            this.explanationLabel.TabIndex = 4;
            this.explanationLabel.Text = resources.GetString("explanationLabel.Text");
            // 
            // allCheck
            // 
            this.allCheck.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.allCheck.AutoSize = true;
            this.allCheck.Location = new System.Drawing.Point(10, 389);
            this.allCheck.Name = "allCheck";
            this.allCheck.Size = new System.Drawing.Size(364, 17);
            this.allCheck.TabIndex = 5;
            this.allCheck.Text = "Remove the same extended properties from next contacts.";
            this.allCheck.UseVisualStyleBackColor = true;
            // 
            // DeleteTooBigPropertiesForm
            // 
            this.AcceptButton = this.bbOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.bbCancel;
            this.ClientSize = new System.Drawing.Size(635, 442);
            this.Controls.Add(this.allCheck);
            this.Controls.Add(this.explanationLabel);
            this.Controls.Add(this.propertiesGrid);
            this.Controls.Add(this.bbCancel);
            this.Controls.Add(this.bbOK);
            this.Font = new System.Drawing.Font("Verdana", 8.25F);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DeleteTooBigPropertiesForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Remove extended properties";
            ((System.ComponentModel.ISupportInitialize)(this.propertiesGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button bbOK;
        private System.Windows.Forms.Button bbCancel;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Selected;
        private System.Windows.Forms.DataGridViewTextBoxColumn Key;
        private System.Windows.Forms.DataGridViewTextBoxColumn Value;
        private System.Windows.Forms.CheckBox allCheck;
        private System.Windows.Forms.DataGridView propertiesGrid;

        public bool removeFromAll
        {
            get { return allCheck.Checked; }
        }

        public void AddExtendedProperty(bool selected, string name, string value)
        {
            propertiesGrid.Rows.Add(selected, name, value);
        }

        public void SortExtendedProperties()
        {
            propertiesGrid.Sort(propertiesGrid.Columns["Key"], ListSortDirection.Ascending);
        }

        private Label explanationLabel;

        public DataGridViewRowCollection extendedPropertiesRows
        {
            get { return propertiesGrid.Rows; }
        }
    }
}