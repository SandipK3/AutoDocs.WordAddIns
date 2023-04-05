namespace AutoDocs.WordAddIns
{
    partial class AutoDocs365TaskPane
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnSaveDocument = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabAdmin = new System.Windows.Forms.TabPage();
            this.tabDataFields = new System.Windows.Forms.TabPage();
            this.tabSearch = new System.Windows.Forms.TabPage();
            this.tabTemplate = new System.Windows.Forms.TabPage();
            this.btnTemplate = new System.Windows.Forms.Button();
            this.btnDataField = new System.Windows.Forms.Button();
            this.btnConditionalContent = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.tabControl1.SuspendLayout();
            this.tabAdmin.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnSaveDocument
            // 
            this.btnSaveDocument.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveDocument.Location = new System.Drawing.Point(15, 91);
            this.btnSaveDocument.Name = "btnSaveDocument";
            this.btnSaveDocument.Size = new System.Drawing.Size(257, 23);
            this.btnSaveDocument.TabIndex = 0;
            this.btnSaveDocument.Text = "Save Document";
            this.btnSaveDocument.UseVisualStyleBackColor = true;
            this.btnSaveDocument.Visible = false;
            this.btnSaveDocument.Click += new System.EventHandler(this.btnSaveDocument_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabAdmin);
            this.tabControl1.Controls.Add(this.tabDataFields);
            this.tabControl1.Controls.Add(this.tabSearch);
            this.tabControl1.Controls.Add(this.tabTemplate);
            this.tabControl1.Location = new System.Drawing.Point(3, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(294, 601);
            this.tabControl1.TabIndex = 1;
            // 
            // tabAdmin
            // 
            this.tabAdmin.Controls.Add(this.dataGridView1);
            this.tabAdmin.Controls.Add(this.btnConditionalContent);
            this.tabAdmin.Controls.Add(this.btnDataField);
            this.tabAdmin.Controls.Add(this.btnTemplate);
            this.tabAdmin.Controls.Add(this.btnSaveDocument);
            this.tabAdmin.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabAdmin.Location = new System.Drawing.Point(4, 22);
            this.tabAdmin.Name = "tabAdmin";
            this.tabAdmin.Padding = new System.Windows.Forms.Padding(3);
            this.tabAdmin.Size = new System.Drawing.Size(286, 575);
            this.tabAdmin.TabIndex = 0;
            this.tabAdmin.Text = "Admin";
            this.tabAdmin.UseVisualStyleBackColor = true;
            // 
            // tabDataFields
            // 
            this.tabDataFields.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabDataFields.Location = new System.Drawing.Point(4, 22);
            this.tabDataFields.Name = "tabDataFields";
            this.tabDataFields.Padding = new System.Windows.Forms.Padding(3);
            this.tabDataFields.Size = new System.Drawing.Size(286, 575);
            this.tabDataFields.TabIndex = 1;
            this.tabDataFields.Text = "DataFields";
            this.tabDataFields.UseVisualStyleBackColor = true;
            // 
            // tabSearch
            // 
            this.tabSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabSearch.Location = new System.Drawing.Point(4, 22);
            this.tabSearch.Name = "tabSearch";
            this.tabSearch.Size = new System.Drawing.Size(286, 575);
            this.tabSearch.TabIndex = 2;
            this.tabSearch.Text = "Search";
            this.tabSearch.UseVisualStyleBackColor = true;
            // 
            // tabTemplate
            // 
            this.tabTemplate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabTemplate.Location = new System.Drawing.Point(4, 22);
            this.tabTemplate.Name = "tabTemplate";
            this.tabTemplate.Size = new System.Drawing.Size(286, 575);
            this.tabTemplate.TabIndex = 3;
            this.tabTemplate.Text = "Template";
            this.tabTemplate.UseVisualStyleBackColor = true;
            // 
            // btnTemplate
            // 
            this.btnTemplate.BackColor = System.Drawing.Color.Green;
            this.btnTemplate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTemplate.ForeColor = System.Drawing.Color.White;
            this.btnTemplate.Location = new System.Drawing.Point(15, 18);
            this.btnTemplate.Name = "btnTemplate";
            this.btnTemplate.Size = new System.Drawing.Size(115, 25);
            this.btnTemplate.TabIndex = 1;
            this.btnTemplate.Text = "Template";
            this.btnTemplate.UseVisualStyleBackColor = false;
            this.btnTemplate.Click += new System.EventHandler(this.btnTemplate_Click);
            // 
            // btnDataField
            // 
            this.btnDataField.BackColor = System.Drawing.Color.RoyalBlue;
            this.btnDataField.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDataField.ForeColor = System.Drawing.Color.White;
            this.btnDataField.Location = new System.Drawing.Point(157, 18);
            this.btnDataField.Name = "btnDataField";
            this.btnDataField.Size = new System.Drawing.Size(115, 25);
            this.btnDataField.TabIndex = 2;
            this.btnDataField.Text = "DataField";
            this.btnDataField.UseVisualStyleBackColor = false;
            this.btnDataField.Click += new System.EventHandler(this.btnDataField_Click);
            // 
            // btnConditionalContent
            // 
            this.btnConditionalContent.BackColor = System.Drawing.Color.RoyalBlue;
            this.btnConditionalContent.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnConditionalContent.ForeColor = System.Drawing.Color.White;
            this.btnConditionalContent.Location = new System.Drawing.Point(15, 60);
            this.btnConditionalContent.Name = "btnConditionalContent";
            this.btnConditionalContent.Size = new System.Drawing.Size(257, 25);
            this.btnConditionalContent.TabIndex = 3;
            this.btnConditionalContent.Text = "Conditional Content";
            this.btnConditionalContent.UseVisualStyleBackColor = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(15, 121);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(257, 328);
            this.dataGridView1.TabIndex = 4;
            // 
            // AutoDocs365TaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Name = "AutoDocs365TaskPane";
            this.Size = new System.Drawing.Size(300, 607);
            this.tabControl1.ResumeLayout(false);
            this.tabAdmin.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnSaveDocument;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabAdmin;
        private System.Windows.Forms.TabPage tabDataFields;
        private System.Windows.Forms.TabPage tabSearch;
        private System.Windows.Forms.TabPage tabTemplate;
        private System.Windows.Forms.Button btnConditionalContent;
        private System.Windows.Forms.Button btnDataField;
        private System.Windows.Forms.Button btnTemplate;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}
