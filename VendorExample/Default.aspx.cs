using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            VendorService.PC_Mobile_WebService WS = new VendorService.PC_Mobile_WebService();
            string LeadID = WS.Mobile_LeadoMat_App_Handler_Vendor(4, "46402", "0", "0", "054 0987654;", "TEST@TEST.COM", "test7", "test77", "01/01/2015", "הערות", 243, "4", "נציג", "UT/VRCK5SuOb6bEb10F6WQ==", "", "").LeadNumber.ToString();
            Response.Write("LeadID: " + LeadID);
        }
    }
}