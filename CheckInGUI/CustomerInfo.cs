using System;
using System.Windows.Forms;

namespace CheckInGUI
{
    public partial class CustomerInfo : Form
    {
        public static string custNameData;
        public static string dateRecData;
        public static string claimNumData;
        public static string custTownData;
        public static bool sameCust = false;
        public CustomerInfo()
        {
            InitializeComponent();
        }

        private void custName_TextChanged(object sender, EventArgs e)
        {
            custNameData = custName.Text;
        }

        private void dateRec_TextChanged(object sender, EventArgs e)
        {
            dateRecData = dateRec.Text;
        }

        private void claimNum_TextChanged(object sender, EventArgs e)
        {
            claimNumData = claimNum.Text;
        }

        private void custTown_TextChanged(object sender, EventArgs e)
        {
            custTownData = custTown.Text;
        }

        private void sameCustomer_Checked(object sender, EventArgs e)
        {
            if (sameCustomer.Checked == true)
            {
                sameCust = true;
            }
            else if (sameCustomer.Checked == false)
            {
                sameCust = false;
            }
        }

        private void CustomerInfo_Load(object sender, EventArgs e)
        {
            ExcelReadoutRefresh();
        }

        private void custName_Enter(object sender, EventArgs e)
        {
            custName.SelectAll();
        }

        private void custTown_Enter(object sender, EventArgs e)
        {
            custTown.SelectAll();
        }

        private void dateRec_Enter(object sender, EventArgs e)
        {
            dateRec.SelectAll();
        }

        private void claimNum_Enter(object sender, EventArgs e)
        {
            claimNum.SelectAll();
        }

        private void custInfoSubmit_Click(object sender, EventArgs e)
        {
            Submit_SaveRecord();
        }
    }
}
