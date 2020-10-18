namespace CheckInGUI
{
    public class ExtgRecord
    {
        public string clModelNum;
        public string clSerialNum;
        public int clYearManu;
        public string clWalkinTakeback;
        public bool clPressurized;
        public string clValve;
        public string clORing1;
        public string clORing2;
        public string clSize;
        public string clChemical;
        public string clExtraParts;
        public string clExtraLabor;
        public string clCollar;
        public string clCT;
        public string clCustomer;
        public string clDate;
        public string clClaim;
        public string clOrder;

        public ExtgRecord(string model, string serial, int year, string wiTB, bool psi)
        {
            clModelNum = model;
            clSerialNum = serial;
            clYearManu = year;
            clWalkinTakeback = wiTB;
            clPressurized = psi;
        }
    }

}
