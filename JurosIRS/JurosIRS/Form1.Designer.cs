namespace JurosIRS
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
            openFileDialog = new OpenFileDialog();
            openFileDialog2 = new OpenFileDialog();
            SuspendLayout();
            // 
            // openFileDialog
            // 
            openFileDialog.FileName = "openFileDialog";
            openFileDialog.Filter = "CSV files (*.csv)|*.csv";
            // 
            // openFileDialog2
            // 
            openFileDialog2.FileName = "openFileDialog2";
            openFileDialog2.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            ClientSize = new Size(61, 18);
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "IRS";
            WindowState = FormWindowState.Minimized;
            Load += Form1_Load;
            ResumeLayout(false);
        }

        #endregion

        private OpenFileDialog openFileDialog;
        private OpenFileDialog openFileDialog2;
        private Label lblResult;
    }
}
