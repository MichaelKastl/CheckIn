using System;
using System.Windows.Forms;

namespace CheckInGUI
{
    public partial class CustDatabase : Form
    {

        public static string file;
        public static bool fileChosen = false;
        public CustDatabase()
        {
            InitializeComponent();
        }

        private void selectDBFileButton_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                file = openFileDialog1.FileName;
                fileChosen = true;

            }
        }

        private void buildDatabaseButton_Click(object sender, EventArgs e)
        {
            DatabaseBuilder();
        }
    }
}
