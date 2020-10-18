using System;
using System.Windows.Forms;

namespace Low_Pressure_Tester
{
    public partial class CalibrationWindow : Form
    {
        public static string retesterInit;
        public static string masterPSI;
        public static string workingPSI;
        public CalibrationWindow()
        {
            InitializeComponent();
        }

        private void RetestersInitBox_TextChanged(object sender, EventArgs e)
        {
            retesterInit = RetestersInitBox.Text;
        }

        private void MasterPressureBox_TextChanged(object sender, EventArgs e)
        {
            masterPSI = MasterPressureBox.Text;
        }

        private void WorkingPressureBox_TextChanged(object sender, EventArgs e)
        {
            workingPSI = WorkingPressureBox.Text;
        }
    }
}
