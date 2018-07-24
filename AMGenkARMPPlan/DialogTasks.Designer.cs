namespace AMGenkARMPPlan
{
    partial class DialogTasks
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
            this.components = new System.ComponentModel.Container();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.lbExcelFile = new System.Windows.Forms.Label();
            this.txtARMPTasksFile = new System.Windows.Forms.TextBox();
            this.btnBrowseT = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnImport = new System.Windows.Forms.Button();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.cbClipboard = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(6, 142);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(251, 32);
            this.btnCancel.TabIndex = 16;
            this.btnCancel.Text = "Annuleren";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(215, 9);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(264, 18);
            this.lblTitle.TabIndex = 1;
            this.lblTitle.Text = "Verversen ARMP Onderhoudsdata";
            this.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbExcelFile
            // 
            this.lbExcelFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbExcelFile.Location = new System.Drawing.Point(3, 59);
            this.lbExcelFile.Name = "lbExcelFile";
            this.lbExcelFile.Size = new System.Drawing.Size(291, 29);
            this.lbExcelFile.TabIndex = 99;
            this.lbExcelFile.Text = "Selecteer de ARMP taken importeer file:";
            this.lbExcelFile.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtARMPTasksFile
            // 
            this.txtARMPTasksFile.Enabled = false;
            this.txtARMPTasksFile.Location = new System.Drawing.Point(300, 64);
            this.txtARMPTasksFile.Name = "txtARMPTasksFile";
            this.txtARMPTasksFile.Size = new System.Drawing.Size(588, 20);
            this.txtARMPTasksFile.TabIndex = 0;
            this.txtARMPTasksFile.WordWrap = false;
            // 
            // btnBrowseT
            // 
            this.btnBrowseT.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowseT.Location = new System.Drawing.Point(894, 52);
            this.btnBrowseT.Name = "btnBrowseT";
            this.btnBrowseT.Size = new System.Drawing.Size(30, 28);
            this.btnBrowseT.TabIndex = 1;
            this.btnBrowseT.Text = "...";
            this.btnBrowseT.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnBrowseT.UseVisualStyleBackColor = true;
            this.btnBrowseT.Click += new System.EventHandler(this.btnBrowseT_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnImport
            // 
            this.btnImport.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImport.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnImport.Location = new System.Drawing.Point(6, 96);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(251, 32);
            this.btnImport.TabIndex = 13;
            this.btnImport.Text = "Importeer";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(291, 29);
            this.label1.TabIndex = 100;
            this.label1.Text = "Importeer rechtstreeks van het clipboard: ";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cbClipboard
            // 
            this.cbClipboard.AutoSize = true;
            this.cbClipboard.Checked = true;
            this.cbClipboard.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbClipboard.Location = new System.Drawing.Point(300, 37);
            this.cbClipboard.Name = "cbClipboard";
            this.cbClipboard.Size = new System.Drawing.Size(15, 14);
            this.cbClipboard.TabIndex = 101;
            this.cbClipboard.UseVisualStyleBackColor = true;
            this.cbClipboard.CheckedChanged += new System.EventHandler(this.cbClipboard_CheckedChanged);
            // 
            // DialogTasks
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(927, 186);
            this.Controls.Add(this.cbClipboard);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.btnBrowseT);
            this.Controls.Add(this.txtARMPTasksFile);
            this.Controls.Add(this.lbExcelFile);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.btnCancel);
            this.Name = "DialogTasks";
            this.Text = "Importeren ARMP naar Planning";
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lbExcelFile;
        private System.Windows.Forms.TextBox txtARMPTasksFile;
        private System.Windows.Forms.Button btnBrowseT;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox cbClipboard;
    }
 }