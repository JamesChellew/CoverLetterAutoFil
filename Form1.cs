using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using System.Security.AccessControl;
using Word = Microsoft.Office.Interop.Word;

namespace CoverLetterAutoFil
{
    public partial class Form1 : Form
    {
        public Word.Application? application;
        public Word.Document? doc;
        public Form1()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents|*.docx";
            string filePath;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
            }
            else
            {
                txtFilePath.Text = "Error";
                return;
            }
            txtFilePath.Text = filePath;
            application = new();
            doc = application.Documents.Open(filePath);
            application.Visible = true;
        }

        private void btnReplaceAndSave_Click(object sender, EventArgs e)
        {
            string filePath = txtFilePath.Text;
            string yourName = txtYourName.Text;
            string yourLastName = txtYourLastName.Text;
            string yourAddress = txtYourAddress.Text;
            string yourEmail = txtYourEmail.Text;
            string yourCity = txtYourCity.Text;
            string yourState = txtYourState.Text;
            string yourPostCode = txtYourPostcode.Text;
            string yourPhoneNumber = txtYourPhoneNumber.Text;
            string date = DateAndTime.Today.Date.ToString();
            string companyName = txtCompanyName.Text;
            string companyPosition = txtCompanyPosition.Text;
            string companyPerson = txtCompanyPerson.Text;

            if (
                string.IsNullOrWhiteSpace(filePath) ||
                string.IsNullOrWhiteSpace(yourName) ||
                string.IsNullOrWhiteSpace(companyName) ||
                string.IsNullOrWhiteSpace(yourAddress) ||
                string.IsNullOrWhiteSpace(yourCity) ||
                string.IsNullOrWhiteSpace(yourState) ||
                string.IsNullOrWhiteSpace(yourPostCode) ||
                string.IsNullOrWhiteSpace(yourPhoneNumber) ||
                string.IsNullOrWhiteSpace(companyName) ||
                string.IsNullOrWhiteSpace(companyPosition))
            {
                MessageBox.Show("Please fill all fields.");
                return;
            }

            if (string.IsNullOrWhiteSpace(companyPerson))
            {
                companyPerson = "whom this reaches";
            }

            if (application == null || doc == null)
            {
                MessageBox.Show("nothing is open yet");
                return;
            }

            ReplacePlaceholder("[Your Name]", yourName);
            ReplacePlaceholder("[Your Last Name]", yourLastName);
            ReplacePlaceholder("[Company Name]", companyName);
            ReplacePlaceholder("[Your Address]", yourAddress);
            ReplacePlaceholder("[City, State, Postcode]", string.Concat(yourCity, " ", yourState, " ", yourPostCode));
            ReplacePlaceholder("[Your City]", yourCity);
            ReplacePlaceholder("[Your State]", yourState);
            ReplacePlaceholder("[Your Postcode]", yourPostCode);
            ReplacePlaceholder("[Your Email Address]", yourEmail);
            ReplacePlaceholder("[Your Phone Number]", yourPhoneNumber);
            ReplacePlaceholder("[Date]", date);
            ReplacePlaceholder("[Hiring Manager's Name]", companyPerson);
            ReplacePlaceholder("[Position Title]", companyPosition);
            ReplacePlaceholder("[Your Full Name]", string.Concat(yourName, " ", yourLastName));
            string newFilePath = filePath.Replace(".docx", "_modified.docx");

            doc.SaveAs2(newFilePath);
            doc.Close();
            application.Quit();
            MessageBox.Show($"File saved as: {newFilePath}");
        }
        private void ReplacePlaceholder(string placeholder, string replacementText)
        {
            if (application == null)
            {
                MessageBox.Show("There is no Document Open");
                return;
            }
            Word.Find findObject = application.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = placeholder;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = replacementText;
            findObject.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }


        private void InsertWordAtCursor(string word)
        {
            // checks if application is running 
            if (application == null)
            {
                return;
            }
            Word.Selection currentSelection = application.Selection;
            // Retain user options
            bool userOvertype = application.Options.Overtype;
            if (application.Options.Overtype)
            {
                application.Options.Overtype = false;
            }
            if (application.Options.Overtype)
            {
                application.Options.Overtype = false;
            }
            // Test to see if selection is an insertion point.
            if (currentSelection.Type == Word.WdSelectionType.wdSelectionIP)
            {
                currentSelection.TypeText(word);
            }
            else if (currentSelection.Type == Word.WdSelectionType.wdSelectionNormal)
            {
                // Move to start of selection.
                if (application.Options.ReplaceSelection)
                {
                    object direction = Word.WdCollapseDirection.wdCollapseStart;
                    currentSelection.Collapse(ref direction);
                }
                currentSelection.TypeText(word);
            }
            application.Options.Overtype = userOvertype;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (application != null)
            {
                application.Quit();
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            InsertWordAtCursor("[Your Name]");
        }

        private void btnLastName_Click(object sender, EventArgs e)
        {
            InsertWordAtCursor("[Your Last Name]");
        }

        private void btnYourEmail_Click(object sender, EventArgs e)
        {
            InsertWordAtCursor("[Your Email Address]");

        }

        private void btnYourPhoneNumber_Click(object sender, EventArgs e)
        {
            InsertWordAtCursor("[Your Phone Number]");

        }

        private void btnYourCity_Click(object sender, EventArgs e)
        {
            InsertWordAtCursor("[Your City]");

        }

        private void btnYourState_Click(object sender, EventArgs e)
        {
            InsertWordAtCursor("[Your State]");

        }

        private void btnYourPostcode_Click(object sender, EventArgs e)
        {
            InsertWordAtCursor("[Your Postcode]");

        }

        private void btnYourAddress_Click(object sender, EventArgs e)
        {
            InsertWordAtCursor("[Your Address]");

        }

        private void btnCompanyName_Click(object sender, EventArgs e)
        {
            InsertWordAtCursor("[Company Name]");

        }

        private void btnCompanyPosition_Click(object sender, EventArgs e)
        {
            InsertWordAtCursor("[Position Title]");

        }

        private void btnCompanyPerson_Click(object sender, EventArgs e)
        {
            InsertWordAtCursor("[Hiring Manager's Name]");

        }
    }
}