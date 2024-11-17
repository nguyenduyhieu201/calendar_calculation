namespace CalendarCalculation
{
    partial class frMain
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
            browserButton = new Button();
            txtOutputFilePath = new TextBox();
            okButton = new Button();
            SuspendLayout();
            // 
            // browserButton
            // 
            browserButton.Location = new Point(652, 23);
            browserButton.Name = "button1";
            browserButton.Size = new Size(94, 29);
            browserButton.TabIndex = 0;
            browserButton.Text = "Browse";
            browserButton.UseVisualStyleBackColor = true;
            browserButton.Click += browserButton_Click;
            // 
            // txtOutputFilePath
            // 
            txtOutputFilePath.Location = new Point(22, 23);
            txtOutputFilePath.Name = "textBox1";
            txtOutputFilePath.Size = new Size(599, 27);
            txtOutputFilePath.TabIndex = 1;
            txtOutputFilePath.TextChanged += textBox1_TextChanged;
            // 
            // okButton
            // 
            okButton.Location = new Point(328, 77);
            okButton.Name = "button2";
            okButton.Size = new Size(116, 29);
            okButton.TabIndex = 2;
            okButton.Text = "Ok";
            okButton.UseVisualStyleBackColor = true;
            okButton.Click += okButton_Click;
            // 
            // frMain
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(763, 130);
            Controls.Add(okButton);
            Controls.Add(txtOutputFilePath);
            Controls.Add(browserButton);
            Name = "frMain";
            Text = "CalendarCalculation";
            Load += frMain_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button browserButton;
        private TextBox txtOutputFilePath;
        private Button okButton;
    }
}
