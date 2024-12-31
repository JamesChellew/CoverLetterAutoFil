namespace CoverLetterAutoFil
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private static FileStream OpenCoverLetter()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Title = "Select a File",
                Filter = "*.docx",
                DefaultExt = "*.docx",
                Multiselect = false,
                CheckFileExists = true,
                CheckPathExists = true,
                RestoreDirectory = true,
                InitialDirectory = Application.StartupPath
            };
            
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.OpenFile();
            }
        }
    }
}