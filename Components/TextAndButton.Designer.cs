namespace CoverLetterAutoFil.Components
{
    partial class TextAndButton
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
            TextBox = new TextBox();
            Button = new Button();
            // 
            // TextBox
            // 
            TextBox.Name = "TextBox";
            TextBox.Size = new Size(100, 27);
            TextBox.TabIndex = 0;
            // 
            // Button
            // 
            Button.Name = "Button";
            Button.Size = new Size(75, 23);
            Button.TabIndex = 0;
            Button.Text = "button1";
            Button.UseVisualStyleBackColor = true;
        }

        #endregion

        private TextBox TextBox;
        private Button Button;
    }
}
