using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OrbProjectDoc
{
    public partial class DocIdProp_Uc : UserControl
    {
        public DocIdProp_Uc()
        {
            InitializeComponent();
        }

        public void UpdateFrom()
        {
            Office.DocumentProperties myCustomProp =
                Globals.ThisDocument.CustomDocumentProperties as Office.DocumentProperties;

            if (OrbHwDocTool.CustomPropertyExist("orbDocCode"))
                this.myDocIdProp_Uc_Wpf.txtOrbDocCode.Text = (string)myCustomProp["orbDocCode"].Value;

            if (OrbHwDocTool.CustomPropertyExist("orbDocTittle"))
                this.myDocIdProp_Uc_Wpf.txtOrbDocTittle.Text = (string)myCustomProp["orbDocTittle"].Value;

            if (OrbHwDocTool.CustomPropertyExist("orbDocShortTittle"))
                this.myDocIdProp_Uc_Wpf.txtOrbDocShortTittle.Text = (string)myCustomProp["orbDocShortTittle"].Value;

            if (OrbHwDocTool.CustomPropertyExist("orbDocContract"))
                this.myDocIdProp_Uc_Wpf.txtOrbDocContract.Text = (string)myCustomProp["orbDocContract"].Value;

            if (OrbHwDocTool.CustomPropertyExist("orbDocContractTittle"))
                this.myDocIdProp_Uc_Wpf.txtOrbDocContractTittle.Text = (string)myCustomProp["orbDocContractTittle"].Value;

            if (OrbHwDocTool.CustomPropertyExist("orbDocProgram"))
                this.myDocIdProp_Uc_Wpf.txtOrbDocProgram.Text = (string)myCustomProp["orbDocProgram"].Value;

            if (OrbHwDocTool.CustomPropertyExist("orbDocProject"))
                this.myDocIdProp_Uc_Wpf.txtOrbDocProject.Text = (string)myCustomProp["orbDocProject"].Value;

            if (OrbHwDocTool.CustomPropertyExist("orbDocIssue"))
                this.myDocIdProp_Uc_Wpf.txtOrbDocIssue.Text = (string)myCustomProp["orbDocIssue"].Value;

            if (OrbHwDocTool.CustomPropertyExist("orbDocIssueDate"))
                this.myDocIdProp_Uc_Wpf.dateOrbDocIssueDate.SelectedDate = (DateTime)myCustomProp["orbDocIssueDate"].Value;

            if (OrbHwDocTool.CustomPropertyExist("orbDocDrlNum"))
                this.myDocIdProp_Uc_Wpf.txtOrbDocDrlNum.Text = Convert.ToString((int)myCustomProp["orbDocDrlNum"].Value);
        }

        public void SaveChange()
        {
            Office.DocumentProperties myCustomProp =
                Globals.ThisDocument.CustomDocumentProperties as Office.DocumentProperties;

            if (OrbHwDocTool.CustomPropertyExist("orbDocCode"))
                myCustomProp["orbDocCode"].Value = myDocIdProp_Uc_Wpf.txtOrbDocCode.Text;

            if (OrbHwDocTool.CustomPropertyExist("orbDocTittle"))
                myCustomProp["orbDocTittle"].Value = myDocIdProp_Uc_Wpf.txtOrbDocTittle.Text;

            if (OrbHwDocTool.CustomPropertyExist("orbDocShortTittle"))
                myCustomProp["orbDocShortTittle"].Value = myDocIdProp_Uc_Wpf.txtOrbDocShortTittle.Text;

            if (OrbHwDocTool.CustomPropertyExist("orbDocContract"))
                myCustomProp["orbDocContract"].Value = myDocIdProp_Uc_Wpf.txtOrbDocContract.Text;

            if (OrbHwDocTool.CustomPropertyExist("orbDocContractTittle"))
                myCustomProp["orbDocContractTittle"].Value = myDocIdProp_Uc_Wpf.txtOrbDocContractTittle.Text;

            if (OrbHwDocTool.CustomPropertyExist("orbDocProgram"))
                myCustomProp["orbDocProgram"].Value = myDocIdProp_Uc_Wpf.txtOrbDocProgram.Text;

            if (OrbHwDocTool.CustomPropertyExist("orbDocProject"))
                myCustomProp["orbDocProject"].Value = myDocIdProp_Uc_Wpf.txtOrbDocProject.Text;

            if (OrbHwDocTool.CustomPropertyExist("orbDocDrlNum"))
                myCustomProp["orbDocDrlNum"].Value = Convert.ToInt32(myDocIdProp_Uc_Wpf.txtOrbDocDrlNum.Text);

            if (OrbHwDocTool.CustomPropertyExist("orbDocIssueDate"))
                myCustomProp["orbDocIssueDate"].Value = myDocIdProp_Uc_Wpf.dateOrbDocIssueDate.SelectedDate;

            OrbHwDocTool.UpdateAllDocFields();
        }
    }
}
