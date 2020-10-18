using System;
using System.Windows.Forms;

namespace CheckInGUI
{
    public partial class NewTemplate : Form
    {

        public static bool templateVerify = false;
        public static string entryModel;
        public static string entryVS;
        public static string entryOR1;
        public static string entryOR2;
        public static string entrySize;
        public static string entryChemical;
        public static string entryExtraParts;
        public static string entryExtraLabor;
        public static int entryHTYearInt;
        public static int entry6YearInt;
        public static bool entryCollar;
        public static bool entryCT;
        public static bool persistPartNumbers;

        public NewTemplate()
        {
            InitializeComponent();
            this.ShowInTaskbar = false;


        }


        private void NewModelNum_TextChanged(object sender, EventArgs e)
        {
            entryModel = NewModelNum.Text;
        }

        private void NewVS_TextChanged(object sender, EventArgs e)
        {
            entryVS = NewVS.Text;
        }

        private void NewOR1_TextChanged(object sender, EventArgs e)
        {
            entryOR1 = NewOR1.Text;
        }

        private void NewOR2_TextChanged(object sender, EventArgs e)
        {
            entryOR2 = NewOR2.Text;
        }

        private void NewSize_TextChanged(object sender, EventArgs e)
        {
            entrySize = NewSize.Text;
        }

        private void NewChemical_TextChanged(object sender, EventArgs e)
        {
            entryChemical = NewChemical.Text;
        }

        private void NewExtraParts_TextChanged(object sender, EventArgs e)
        {
            entryExtraParts = NewExtraParts.Text;
        }

        private void NewExtraLabor_TextChanged(object sender, EventArgs e)
        {
            entryExtraLabor = NewExtraLabor.Text;
        }

        private void NewHTYear_Leave(object sender, EventArgs e)
        {
            string entryHTYear = NewHTYear.Text;
            bool tryParse = Int32.TryParse(entryHTYear, out int htYearParse);
            if (tryParse == false)
            {
                MessageBox.Show("Please enter a valid number in the 'New HT Year' field.", "Invalid Input",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
                NewHTYear.Focus();
            }
            else if (tryParse == true)
            {
                entryHTYearInt = htYearParse;
            }

        }

        private void New6YRYear_Leave(object sender, EventArgs e)
        {
            string entry6YRYear = New6YRYear.Text;
            bool tryParse = Int32.TryParse(entry6YRYear, out int sixYearParse);
            if (tryParse == false)
            {
                MessageBox.Show("Please enter a valid number in the 'New 6YR Year' field.", "Invalid Input",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
                New6YRYear.Focus();
            }
            else if (tryParse == true)
            {
                entry6YearInt = sixYearParse;
            }
        }

        private void NewCollarYesB_CheckedChanged(object sender, EventArgs e)
        {
            entryCollar = true;
        }

        private void NewCollarNoB_CheckedChanged(object sender, EventArgs e)
        {
            entryCollar = false;
        }

        private void NewCTYesB_CheckedChanged(object sender, EventArgs e)
        {
            entryCT = true;
        }

        private void NewCTNoB_CheckedChanged(object sender, EventArgs e)
        {
            entryCT = false;
        }

        public void SubmissionB_Click(object sender, EventArgs e)
        {

            TemplateBuilder();
            Control_ReturnFocus();

        }

        private void NewModelNum_Enter(object sender, EventArgs e)
        {
            NewModelNum.SelectAll();
        }

        private void NewVS_Enter(object sender, EventArgs e)
        {
            NewVS.SelectAll();
        }

        private void NewOR1_Enter(object sender, EventArgs e)
        {
            NewOR1.SelectAll();
        }

        private void NewOR2_Enter(object sender, EventArgs e)
        {
            NewOR2.SelectAll();
        }

        private void NewSize_Enter(object sender, EventArgs e)
        {
            NewSize.SelectAll();
        }

        private void NewChemical_Enter(object sender, EventArgs e)
        {
            NewChemical.SelectAll();
        }

        private void NewExtraParts_Enter(object sender, EventArgs e)
        {
            NewExtraParts.SelectAll();
        }

        private void NewExtraLabor_Enter(object sender, EventArgs e)
        {
            NewExtraLabor.SelectAll();
        }

        private void NewHTYear_Enter(object sender, EventArgs e)
        {
            NewHTYear.SelectAll();
        }

        private void New6YRYear_Enter(object sender, EventArgs e)
        {
            New6YRYear.SelectAll();
        }

        private void multiSizeBox_CheckedChanged(object sender, EventArgs e)
        {
            if (multiSizeBox.Checked == true)
            {
                persistPartNumbers = true;
            }
            else
            {
                persistPartNumbers = false;
            }
        }
    }
}
