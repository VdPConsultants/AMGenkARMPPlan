namespace AMGenkARMPPlan
{
    partial class DialogFilter
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
            this.lblKolommen = new System.Windows.Forms.Label();
            this.btnAnuleren = new System.Windows.Forms.Button();
            this.btnFilteren = new System.Windows.Forms.Button();
            this.cbSelectAll = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // lblKolommen
            // 
            this.lblKolommen.AutoSize = true;
            this.lblKolommen.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblKolommen.Location = new System.Drawing.Point(8, 9);
            this.lblKolommen.Name = "lblKolommen";
            this.lblKolommen.Size = new System.Drawing.Size(115, 20);
            this.lblKolommen.TabIndex = 1;
            this.lblKolommen.Text = "Medewerkers";
            // 
            // btnAnuleren
            // 
            this.btnAnuleren.Location = new System.Drawing.Point(348, 12);
            this.btnAnuleren.Name = "btnAnuleren";
            this.btnAnuleren.Size = new System.Drawing.Size(143, 34);
            this.btnAnuleren.TabIndex = 2;
            this.btnAnuleren.Text = "Anuleren";
            this.btnAnuleren.UseVisualStyleBackColor = true;
            this.btnAnuleren.Click += new System.EventHandler(this.btnAnuleren_Click);
            // 
            // btnFilteren
            // 
            this.btnFilteren.Location = new System.Drawing.Point(348, 52);
            this.btnFilteren.Name = "btnFilteren";
            this.btnFilteren.Size = new System.Drawing.Size(143, 34);
            this.btnFilteren.TabIndex = 3;
            this.btnFilteren.Text = "Filteren";
            this.btnFilteren.UseVisualStyleBackColor = true;
            this.btnFilteren.Click += new System.EventHandler(this.btnFilteren_Click);
            // 
            // cbSelectAll
            // 
            this.cbSelectAll.AutoSize = true;
            this.cbSelectAll.Location = new System.Drawing.Point(10, 40);
            this.cbSelectAll.Name = "cbSelectAll";
            this.cbSelectAll.Size = new System.Drawing.Size(49, 17);
            this.cbSelectAll.TabIndex = 5;
            this.cbSelectAll.Text = "(Alle)";
            this.cbSelectAll.UseVisualStyleBackColor = true;
            this.cbSelectAll.CheckedChanged += new System.EventHandler(this.cbSelectAll_CheckedChanged);
            // 
            // DialogFilter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(503, 533);
            this.Controls.Add(this.cbSelectAll);
            this.Controls.Add(this.btnFilteren);
            this.Controls.Add(this.btnAnuleren);
            this.Controls.Add(this.lblKolommen);
            this.Name = "DialogFilter";
            this.Text = "DialogFilter";
            this.Load += new System.EventHandler(this.DialogFilter_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label lblKolommen;
        private System.Windows.Forms.Button btnAnuleren;
        private System.Windows.Forms.Button btnFilteren;
        private System.Windows.Forms.CheckBox cbSelectAll;
    }
}