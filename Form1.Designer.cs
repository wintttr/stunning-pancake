namespace WinFormsApp1
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.loadFileButton = new System.Windows.Forms.Button();
            this.disciplineCheckedList = new System.Windows.Forms.CheckedListBox();
            this.generateOutputButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // loadFileButton
            // 
            this.loadFileButton.Location = new System.Drawing.Point(11, 12);
            this.loadFileButton.Name = "loadFileButton";
            this.loadFileButton.Size = new System.Drawing.Size(94, 51);
            this.loadFileButton.TabIndex = 0;
            this.loadFileButton.Text = "Загрузить файл";
            this.loadFileButton.UseVisualStyleBackColor = true;
            this.loadFileButton.Click += new System.EventHandler(this.loadFileButton_Click);
            // 
            // disciplineCheckedList
            // 
            this.disciplineCheckedList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.disciplineCheckedList.FormattingEnabled = true;
            this.disciplineCheckedList.Location = new System.Drawing.Point(112, 35);
            this.disciplineCheckedList.Name = "disciplineCheckedList";
            this.disciplineCheckedList.Size = new System.Drawing.Size(398, 290);
            this.disciplineCheckedList.TabIndex = 1;
            // 
            // generateOutputButton
            // 
            this.generateOutputButton.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.generateOutputButton.Location = new System.Drawing.Point(0, 331);
            this.generateOutputButton.Name = "generateOutputButton";
            this.generateOutputButton.Size = new System.Drawing.Size(522, 42);
            this.generateOutputButton.TabIndex = 2;
            this.generateOutputButton.Text = "Генерация";
            this.generateOutputButton.UseVisualStyleBackColor = true;
            this.generateOutputButton.Click += new System.EventHandler(this.generateOutputButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(112, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(140, 20);
            this.label1.TabIndex = 3;
            this.label1.Text = "Список дисциплин";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(522, 373);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.generateOutputButton);
            this.Controls.Add(this.disciplineCheckedList);
            this.Controls.Add(this.loadFileButton);
            this.MinimumSize = new System.Drawing.Size(540, 420);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button loadFileButton;
        private CheckedListBox disciplineCheckedList;
        private Button generateOutputButton;
        private Label label1;
    }
}