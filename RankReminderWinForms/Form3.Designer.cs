
namespace RankReminderWinForms
{
    partial class Form3
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form3));
            this.Button_CloseSettings = new System.Windows.Forms.Button();
            this.Label_OperationsWithDB = new System.Windows.Forms.Label();
            this.Button_LoadDB = new System.Windows.Forms.Button();
            this.Button_UnloadDB = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.Button_RecreateDB = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Button_CloseSettings
            // 
            this.Button_CloseSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Button_CloseSettings.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Button_CloseSettings.Image = ((System.Drawing.Image)(resources.GetObject("Button_CloseSettings.Image")));
            this.Button_CloseSettings.Location = new System.Drawing.Point(498, 356);
            this.Button_CloseSettings.Name = "Button_CloseSettings";
            this.Button_CloseSettings.Size = new System.Drawing.Size(170, 60);
            this.Button_CloseSettings.TabIndex = 0;
            this.Button_CloseSettings.Text = "Закрыть";
            this.Button_CloseSettings.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.Button_CloseSettings.UseVisualStyleBackColor = true;
            this.Button_CloseSettings.Click += new System.EventHandler(this.Button_CloseSettings_Click);
            // 
            // Label_OperationsWithDB
            // 
            this.Label_OperationsWithDB.AutoSize = true;
            this.Label_OperationsWithDB.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Label_OperationsWithDB.Location = new System.Drawing.Point(6, 20);
            this.Label_OperationsWithDB.Name = "Label_OperationsWithDB";
            this.Label_OperationsWithDB.Size = new System.Drawing.Size(232, 20);
            this.Label_OperationsWithDB.TabIndex = 1;
            this.Label_OperationsWithDB.Text = "Операции с базой данных:";
            // 
            // Button_LoadDB
            // 
            this.Button_LoadDB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Button_LoadDB.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Button_LoadDB.Image = ((System.Drawing.Image)(resources.GetObject("Button_LoadDB.Image")));
            this.Button_LoadDB.Location = new System.Drawing.Point(228, 52);
            this.Button_LoadDB.Name = "Button_LoadDB";
            this.Button_LoadDB.Size = new System.Drawing.Size(200, 60);
            this.Button_LoadDB.TabIndex = 0;
            this.Button_LoadDB.Text = "Подрузить БД";
            this.Button_LoadDB.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.Button_LoadDB.UseVisualStyleBackColor = true;
            this.Button_LoadDB.Click += new System.EventHandler(this.Button_LoadDB_Click);
            // 
            // Button_UnloadDB
            // 
            this.Button_UnloadDB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Button_UnloadDB.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Button_UnloadDB.Image = ((System.Drawing.Image)(resources.GetObject("Button_UnloadDB.Image")));
            this.Button_UnloadDB.Location = new System.Drawing.Point(10, 52);
            this.Button_UnloadDB.Name = "Button_UnloadDB";
            this.Button_UnloadDB.Size = new System.Drawing.Size(200, 60);
            this.Button_UnloadDB.TabIndex = 0;
            this.Button_UnloadDB.Text = "Выгрузить БД";
            this.Button_UnloadDB.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.Button_UnloadDB.UseVisualStyleBackColor = true;
            this.Button_UnloadDB.Click += new System.EventHandler(this.Button_UnloadDB_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.Button_RecreateDB);
            this.groupBox1.Controls.Add(this.Label_OperationsWithDB);
            this.groupBox1.Controls.Add(this.Button_LoadDB);
            this.groupBox1.Controls.Add(this.Button_UnloadDB);
            this.groupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(656, 131);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            // 
            // Button_RecreateDB
            // 
            this.Button_RecreateDB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Button_RecreateDB.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Button_RecreateDB.Image = ((System.Drawing.Image)(resources.GetObject("Button_RecreateDB.Image")));
            this.Button_RecreateDB.Location = new System.Drawing.Point(446, 52);
            this.Button_RecreateDB.Name = "Button_RecreateDB";
            this.Button_RecreateDB.Size = new System.Drawing.Size(200, 60);
            this.Button_RecreateDB.TabIndex = 2;
            this.Button_RecreateDB.Text = "Пересоздать БД";
            this.Button_RecreateDB.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.Button_RecreateDB.UseVisualStyleBackColor = true;
            this.Button_RecreateDB.Click += new System.EventHandler(this.Button_RecreateDB_Click);
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(680, 450);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.Button_CloseSettings);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form3";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Настройки программы";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button Button_CloseSettings;
        private System.Windows.Forms.Label Label_OperationsWithDB;
        private System.Windows.Forms.Button Button_LoadDB;
        private System.Windows.Forms.Button Button_UnloadDB;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button Button_RecreateDB;
    }
}