namespace KingStoneModify
{
    partial class SelectHistoryData_No1
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
            this.dataGridView_data = new System.Windows.Forms.DataGridView();
            this.button_Sure = new System.Windows.Forms.Button();
            this.Grp_Display = new System.Windows.Forms.GroupBox();
            this.button_Return = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_data)).BeginInit();
            this.Grp_Display.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView_data
            // 
            this.dataGridView_data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_data.Location = new System.Drawing.Point(3, 17);
            this.dataGridView_data.Name = "dataGridView_data";
            this.dataGridView_data.RowTemplate.Height = 23;
            this.dataGridView_data.Size = new System.Drawing.Size(638, 198);
            this.dataGridView_data.TabIndex = 0;
            this.dataGridView_data.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_data_CellClick);
            // 
            // button_Sure
            // 
            this.button_Sure.Location = new System.Drawing.Point(491, 235);
            this.button_Sure.Name = "button_Sure";
            this.button_Sure.Size = new System.Drawing.Size(75, 23);
            this.button_Sure.TabIndex = 7;
            this.button_Sure.Text = "确定";
            this.button_Sure.UseVisualStyleBackColor = true;
            this.button_Sure.Click += new System.EventHandler(this.button_Sure_Click);
            // 
            // Grp_Display
            // 
            this.Grp_Display.Controls.Add(this.dataGridView_data);
            this.Grp_Display.Location = new System.Drawing.Point(17, 4);
            this.Grp_Display.Name = "Grp_Display";
            this.Grp_Display.Size = new System.Drawing.Size(651, 225);
            this.Grp_Display.TabIndex = 6;
            this.Grp_Display.TabStop = false;
            this.Grp_Display.Text = "数据库";
            // 
            // button_Return
            // 
            this.button_Return.Location = new System.Drawing.Point(593, 235);
            this.button_Return.Name = "button_Return";
            this.button_Return.Size = new System.Drawing.Size(75, 23);
            this.button_Return.TabIndex = 8;
            this.button_Return.Text = "返回";
            this.button_Return.UseVisualStyleBackColor = true;
            this.button_Return.Click += new System.EventHandler(this.button_Return_Click);
            // 
            // SelectHistoryData_No1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(684, 262);
            this.Controls.Add(this.button_Sure);
            this.Controls.Add(this.Grp_Display);
            this.Controls.Add(this.button_Return);
            this.Name = "SelectHistoryData_No1";
            this.Text = "SelectHistoryData_No1";
            this.Load += new System.EventHandler(this.SelectHistoryData_No1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_data)).EndInit();
            this.Grp_Display.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView_data;
        private System.Windows.Forms.Button button_Sure;
        private System.Windows.Forms.GroupBox Grp_Display;
        private System.Windows.Forms.Button button_Return;
    }
}