namespace AMGenkARMPPlan
{
    partial class Dialog
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
            this.lblVersionHeader = new System.Windows.Forms.Label();
            this.lblAppVersion = new System.Windows.Forms.Label();
            this.lblPublishedVersion = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnBrowseR = new System.Windows.Forms.Button();
            this.txtARMPResourcesFile = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.mcARMPweek = new System.Windows.Forms.MonthCalendar();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(6, 320);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(251, 32);
            this.btnCancel.TabIndex = 16;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(215, 9);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(261, 18);
            this.lblTitle.TabIndex = 1;
            this.lblTitle.Text = "Importeer ARMP Onderhoudsdata";
            this.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbExcelFile
            // 
            this.lbExcelFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbExcelFile.Location = new System.Drawing.Point(3, 237);
            this.lbExcelFile.Name = "lbExcelFile";
            this.lbExcelFile.Size = new System.Drawing.Size(291, 29);
            this.lbExcelFile.TabIndex = 99;
            this.lbExcelFile.Text = "3. Selecteer de ARMP taken importeer file:";
            this.lbExcelFile.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtARMPTasksFile
            // 
            this.txtARMPTasksFile.Location = new System.Drawing.Point(300, 242);
            this.txtARMPTasksFile.Name = "txtARMPTasksFile";
            this.txtARMPTasksFile.Size = new System.Drawing.Size(588, 20);
            this.txtARMPTasksFile.TabIndex = 0;
            this.txtARMPTasksFile.WordWrap = false;
            // 
            // btnBrowseT
            // 
            this.btnBrowseT.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowseT.Location = new System.Drawing.Point(894, 238);
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
            this.btnImport.Location = new System.Drawing.Point(6, 273);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(251, 32);
            this.btnImport.TabIndex = 13;
            this.btnImport.Text = "Importeer";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // lblVersionHeader
            // 
            this.lblVersionHeader.AutoSize = true;
            this.lblVersionHeader.Location = new System.Drawing.Point(295, 283);
            this.lblVersionHeader.Name = "lblVersionHeader";
            this.lblVersionHeader.Size = new System.Drawing.Size(108, 13);
            this.lblVersionHeader.TabIndex = 102;
            this.lblVersionHeader.Text = "ImportProject Add-in: ";
            // 
            // lblAppVersion
            // 
            this.lblAppVersion.AutoSize = true;
            this.lblAppVersion.Location = new System.Drawing.Point(305, 168);
            this.lblAppVersion.Name = "lblAppVersion";
            this.lblAppVersion.Size = new System.Drawing.Size(106, 13);
            this.lblAppVersion.TabIndex = 103;
            this.lblAppVersion.Text = "Application Version:  ";
            // 
            // lblPublishedVersion
            // 
            this.lblPublishedVersion.AutoSize = true;
            this.lblPublishedVersion.Location = new System.Drawing.Point(305, 184);
            this.lblPublishedVersion.Name = "lblPublishedVersion";
            this.lblPublishedVersion.Size = new System.Drawing.Size(152, 13);
            this.lblPublishedVersion.TabIndex = 104;
            this.lblPublishedVersion.Text = "ClickOnce Published Version:  ";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 208);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(286, 29);
            this.label1.TabIndex = 105;
            this.label1.Text = "2. Selecteer de ARMP werknemers importeer file:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnBrowseR
            // 
            this.btnBrowseR.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowseR.Location = new System.Drawing.Point(894, 208);
            this.btnBrowseR.Name = "btnBrowseR";
            this.btnBrowseR.Size = new System.Drawing.Size(30, 28);
            this.btnBrowseR.TabIndex = 107;
            this.btnBrowseR.Text = "...";
            this.btnBrowseR.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnBrowseR.UseVisualStyleBackColor = true;
            this.btnBrowseR.Click += new System.EventHandler(this.btnBrowseR_Click);
            // 
            // txtARMPResourcesFile
            // 
            this.txtARMPResourcesFile.Location = new System.Drawing.Point(301, 214);
            this.txtARMPResourcesFile.Name = "txtARMPResourcesFile";
            this.txtARMPResourcesFile.Size = new System.Drawing.Size(587, 20);
            this.txtARMPResourcesFile.TabIndex = 106;
            this.txtARMPResourcesFile.WordWrap = false;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(3, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(275, 29);
            this.label2.TabIndex = 108;
            this.label2.Text = "1. Selecteer week:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // mcARMPweek
            // 
            this.mcARMPweek.Location = new System.Drawing.Point(301, 46);
            this.mcARMPweek.Name = "mcARMPweek";
            this.mcARMPweek.ShowWeekNumbers = true;
            this.mcARMPweek.TabIndex = 109;
            this.mcARMPweek.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.mcARMPweek_DateChanged);
            // 
            // Dialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(927, 490);
            this.Controls.Add(this.mcARMPweek);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnBrowseR);
            this.Controls.Add(this.txtARMPResourcesFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblPublishedVersion);
            this.Controls.Add(this.lblAppVersion);
            this.Controls.Add(this.lblVersionHeader);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.btnBrowseT);
            this.Controls.Add(this.txtARMPTasksFile);
            this.Controls.Add(this.lbExcelFile);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.btnCancel);
            this.Name = "Dialog";
            this.Text = "Import From Excel";
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
        private System.Windows.Forms.Label lblVersionHeader;
        private System.Windows.Forms.Label lblAppVersion;
        private System.Windows.Forms.Label lblPublishedVersion;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnBrowseR;
        private System.Windows.Forms.TextBox txtARMPResourcesFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MonthCalendar mcARMPweek;
    }
 }