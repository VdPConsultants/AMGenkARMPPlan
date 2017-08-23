namespace AMGenkARMPPlan
{
    partial class DialogATasks
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
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(6, 134);
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
            this.lbExcelFile.Location = new System.Drawing.Point(3, 51);
            this.lbExcelFile.Name = "lbExcelFile";
            this.lbExcelFile.Size = new System.Drawing.Size(291, 29);
            this.lbExcelFile.TabIndex = 99;
            this.lbExcelFile.Text = "Selecteer de ARMP A-taken importeer file:";
            this.lbExcelFile.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtARMPTasksFile
            // 
            this.txtARMPTasksFile.Location = new System.Drawing.Point(300, 56);
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
            this.btnImport.Location = new System.Drawing.Point(6, 88);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(251, 32);
            this.btnImport.TabIndex = 13;
            this.btnImport.Text = "Importeer";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // DialogATasks
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(927, 490);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.btnBrowseT);
            this.Controls.Add(this.txtARMPTasksFile);
            this.Controls.Add(this.lbExcelFile);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.btnCancel);
            this.Name = "DialogATasks";
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
    }
 }