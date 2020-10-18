using System;
using System.Windows.Forms;

namespace Low_Pressure_Tester
{
    public partial class TestingWindow : Form
    {
        public static string testPSI;
        public static string visual;
        public static string dispCode;
        public static string retesterInit;
        public static string manufacturer;
        public static string DOTSpec;
        public TestingWindow()
        {
            InitializeComponent();
            if (EnableManuAndDOTCheck.Checked == true)
            {
                ManufacturerBox.ReadOnly = false;
                DOTSpecBox.ReadOnly = false;
            }
            else
            {
                ManufacturerBox.ReadOnly = true;
                DOTSpecBox.ReadOnly = true;
            }
        }

        private void EnableManuAndDOTCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (EnableManuAndDOTCheck.Checked == true)
            {
                ManufacturerBox.ReadOnly = false;
                DOTSpecBox.ReadOnly = false;
            }
            else
            {
                ManufacturerBox.ReadOnly = true;
                DOTSpecBox.ReadOnly = true;
            }
        }

        private void TestPSIBox_TextChanged(object sender, EventArgs e)
        {
            testPSI = TestPSIBox.Text;
        }

        private void VisualBox_TextChanged(object sender, EventArgs e)
        {
            visual = VisualBox.Text;
        }

        private void DispCodeBox_TextChanged(object sender, EventArgs e)
        {
            dispCode = DispCodeBox.Text;
        }

        private void RetesterInitBox_TextChanged(object sender, EventArgs e)
        {
            retesterInit = RetesterInitBox.Text;
        }

        private void ManufacturerBox_TextChanged(object sender, EventArgs e)
        {
            manufacturer = ManufacturerBox.Text;
        }

        private void DOTSpecBox_TextChanged(object sender, EventArgs e)
        {
            DOTSpec = DOTSpecBox.Text;
        }
    }
}
