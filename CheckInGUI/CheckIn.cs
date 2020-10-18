using System;
using System.Windows.Forms;
namespace CheckInGUI
{
    public partial class CheckIn : Form
    {
        public static string recordModelNum = "blank";
        public static string recordSerialNum;
        public static string recordYearManuStr;
        public static string recordWITB;
        public static bool recordPressurized;
        public static string recordPressurizedStr;
        public static bool persistExtgInfo;
        private void ShowMyOwnedForm2()
        {
            NewTemplate secondForm = new NewTemplate
            {
                Owner = this
            };
            secondForm.Show();
        }
        private void ShowMyOwnedForm3()
        {
            CustomerInfo thirdForm = new CustomerInfo
            {
                Owner = this
            };
            thirdForm.Show();
        }
        private void ShowMyOwnedForm4()
        {
            CustDatabase fourthForm = new CustDatabase
            {
                Owner = this
            };
            fourthForm.Show();
        }
        public CheckIn()
        {
            InitializeComponent();

        }

        private void Button3_Click(object sender, EventArgs e)
        {

        }

        private void Button1_Click_1(object sender, EventArgs e)
        {

        }

        private void Button2_Click_1(object sender, EventArgs e)
        {

        }

        private void ModelNum_TextChanged(object sender, EventArgs e)
        {
            recordModelNum = modelNum.Text;

        }

        private void serialNum_TextChanged(object sender, EventArgs e)
        {
            recordSerialNum = serialNum.Text.ToUpper();

        }

        private void yearManu_TextChanged(object sender, EventArgs e)
        {
            recordYearManuStr = yearManu.Text;

        }

        private void walkinTakeback_TextChanged(object sender, EventArgs e)
        {
            recordWITB = walkinTakeback.Text;

        }

        private void button4_Click(object sender, EventArgs e)
        {
            ShowMyOwnedForm2();
        }

        public void Verify_Click(object sender, EventArgs e)
        {

            NewTemplate.templateVerify = true;

        }

        private void pressurized_TextChanged(object sender, EventArgs e)
        {
            recordPressurizedStr = pressurized.Text;
            string recordPressurizedCAPS = recordPressurizedStr.ToUpper();
            if (recordPressurizedCAPS == "YES")
            {
                recordPressurized = true;
            }
            else
            {
                recordPressurized = false;
            }

        }

        private void custInfoButton_Click(object sender, EventArgs e)
        {
            ShowMyOwnedForm3();
        }

        private void CheckIn_Shown(object sender, EventArgs e)
        {
            ExcelReadoutRefresh();
            this.ActiveControl = modelNum;
            this.indivExLabor.Enabled = false;
            this.indivExParts.Enabled = false;
            BeginningFocus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReadoutRefresh();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenDoc();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            CalculateSheet();
        }

        private void modelNum_Enter(object sender, EventArgs e)
        {
            modelNum.SelectAll();
        }

        private void serialNum_Enter(object sender, EventArgs e)
        {
            serialNum.SelectAll();
        }

#pragma warning disable IDE1006 // Naming Styles
        private void yearManu_Enter(object sender, EventArgs e)
#pragma warning restore IDE1006 // Naming Styles
        {
            yearManu.SelectAll();
        }

        private void walkinTakeback_Enter(object sender, EventArgs e)
        {
            walkinTakeback.SelectAll();
        }

        private void pressurized_Enter(object sender, EventArgs e)
        {
            pressurized.SelectAll();
        }

        private void indivExParts_TextChanged(object sender, EventArgs e)
        {
            globalExtraParts = indivExParts.Text;
            globalExPartsBool = true;
        }

        private void indivExLabor_TextChanged(object sender, EventArgs e)
        {
            globalExtraLabor = indivExLabor.Text;
            globalExLaborBool = true;
        }

        private void extraPartsCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (extraPartsCheck.Checked == true)
            {
                indivExParts.Enabled = true;
            }
            else
            {
                indivExParts.Enabled = false;
            }
        }

        private void extraLaborCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (extraLaborCheck.Checked == true)
            {
                indivExLabor.Enabled = true;
            }
            else
            {
                indivExLabor.Enabled = false;
            }
        }

        private void checkInSubmit_Click(object sender, EventArgs e)
        {
            SubmitButton();
        }

        private void condemnedCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (condemnedCheck.Checked == true)
            {
                globalExtraLabor = "Condemned";
                globalExLaborBool = true;
            }

        }

        private void persistInfoBox_CheckedChanged(object sender, EventArgs e)
        {
            if (persistInfoBox.Checked == true)
            {
                persistExtgInfo = true;
            }
            else
            {
                persistExtgInfo = false;
            }
        }

        private void custDatabaseScreenButton_Click(object sender, EventArgs e)
        {
            ShowMyOwnedForm4();
        }
    }
}
