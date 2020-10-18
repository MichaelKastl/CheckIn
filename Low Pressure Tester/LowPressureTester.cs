using System;
using System.Windows.Forms;

namespace Low_Pressure_Tester
{
    public partial class LowPressureTester : Form
    {
        public static string modelNum;
        public static string serialNum;
        public static string customer;
        public static string custTown;
        public static bool manualCustBool;



        public LowPressureTester()
        {
            InitializeComponent();
            if (manualCustDataCheck.Checked == true)
            {
                manualCustBool = true;
                custNameBox.ReadOnly = false;
                custTownBox.ReadOnly = false;
            }
            else
            {
                manualCustBool = false;
                custNameBox.ReadOnly = true;
                custTownBox.ReadOnly = true;
            }

        }
        public void ShowTestForm()
        {
            TestingWindow testingWindow = new TestingWindow
            {
                Owner = this
            };
            testingWindow.Show();
        }

        public void ShowCalibrationForm()
        {
            CalibrationWindow calibrationWindow = new CalibrationWindow
            {
                Owner = this
            };
            calibrationWindow.Show();
        }

        private void custTownBox_TextChanged(object sender, EventArgs e)
        {
            if (manualCustDataCheck.Checked == true)
            {
                custTown = custTownBox.Text;
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void custNameBox_TextChanged(object sender, EventArgs e)
        {
            if (manualCustDataCheck.Checked == true)
            {
                customer = custNameBox.Text;
            }
        }

        private void serialNumBox_TextChanged(object sender, EventArgs e)
        {
            serialNum = serialNumBox.Text;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void modelNumBox_TextChanged(object sender, EventArgs e)
        {
            modelNum = modelNumBox.Text;
        }

        private void testModeButton_Click(object sender, EventArgs e)
        {
            ShowTestForm();
        }

        private void manualCustDataCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (manualCustDataCheck.Checked == true)
            {
                manualCustBool = true;
                custNameBox.ReadOnly = false;
                custTownBox.ReadOnly = false;
            }
            else
            {
                manualCustBool = false;
                custNameBox.ReadOnly = true;
                custTownBox.ReadOnly = true;
            }
        }

        private void calibrationModeButton_Click(object sender, EventArgs e)
        {
            ShowCalibrationForm();
        }
        private void modelNumBox_Enter(object sender, EventArgs e)
        {
            modelNumBox.SelectAll();
        }
        private void serialNumBox_Enter(object sender, EventArgs e)
        {
            serialNumBox.SelectAll();
        }
        private void custNameBox_Enter(object sender, EventArgs e)
        {
            custNameBox.SelectAll();
        }
        private void custTownBox_Enter(object sender, EventArgs e)
        {
            custTownBox.SelectAll();
        }
    }
}
