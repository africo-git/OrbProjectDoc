using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;


namespace OrbProjectDoc
{
    public static class OrbHwDocTool
    {
        public static void DocActionTaskPaneIni()
        {
            // Carga los controles en el ActionsPane
            Globals.ThisDocument.ActionsPane.Controls.Add(Globals.ThisDocument.myDocIdProp_Uc);
            //Globals.ThisDocument.ActionsPane.Controls.Add(Globals.ThisDocument.myDocVerCtrl_Uc);

            // Almacena los índices de los controles cargados en el ActionsPane
            int myDocIdProp_Uc_index = Globals.ThisDocument.ActionsPane.Controls.GetChildIndex(Globals.ThisDocument.myDocIdProp_Uc);
            //int myDocVerCtrl_Uc_index = Globals.ThisDocument.ActionsPane.Controls.GetChildIndex(Globals.ThisDocument.myDocVerCtrl_Uc);

            // Oculta todos los controles en el ActionsPane
            Globals.ThisDocument.ActionsPane.Controls[myDocIdProp_Uc_index].Visible = false;
            //Globals.ThisDocument.ActionsPane.Controls[myDocVerCtrl_Uc_index].Visible = false;

            // Oculta el Document Actions Task Pane
            Globals.ThisDocument.Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneDocumentActions].Visible = false;
        }

        public static void RestoreFundamentalProp()
        {
            // Se utilizará para acceder a las propiedades personalizadas del documento
            Office.DocumentProperties toolDocCustomProps =
                Globals.ThisDocument.CustomDocumentProperties as Office.DocumentProperties;

            int lastIssueNum;       // Contendrá el número identificativo de la versión actual (cero para la inicial)

            #region PROPIEDADES GENERALES DE IDENTIFICACIÓN
            if (!OrbHwDocTool.CustomPropertyExist("orbCompany"))
                OrbHwDocTool.NewDocCustomProperty("orbCompany", Office.MsoDocProperties.msoPropertyTypeString, "Orbital Sistemas Aeroespaciales, S.L.");

            if (!OrbHwDocTool.CustomPropertyExist("orbCompanyAddress1"))
                OrbHwDocTool.NewDocCustomProperty("orbCompanyAddress1", Office.MsoDocProperties.msoPropertyTypeString, "Carretera de Artica 29, 3ª Planta");

            if (!OrbHwDocTool.CustomPropertyExist("orbCompanyAddress2"))
                OrbHwDocTool.NewDocCustomProperty("orbCompanyAddress2", Office.MsoDocProperties.msoPropertyTypeString, "31013 Artica, Navarra");

            if (!OrbHwDocTool.CustomPropertyExist("orbCompanyAddress3"))
                OrbHwDocTool.NewDocCustomProperty("orbCompanyAddress3", Office.MsoDocProperties.msoPropertyTypeString, "SPAIN");

            if (!OrbHwDocTool.CustomPropertyExist("orbCif"))
                OrbHwDocTool.NewDocCustomProperty("orbCif", Office.MsoDocProperties.msoPropertyTypeString, "CIF: B31954506");

            if (!OrbHwDocTool.CustomPropertyExist("orbDocCode"))
                OrbHwDocTool.NewDocCustomProperty("orbDocCode", Office.MsoDocProperties.msoPropertyTypeString, "Document Code");

            if (!OrbHwDocTool.CustomPropertyExist("orbDocTittle"))
                OrbHwDocTool.NewDocCustomProperty("orbDocTittle", Office.MsoDocProperties.msoPropertyTypeString, "Document Tittle");

            if (!OrbHwDocTool.CustomPropertyExist("orbDocShortTittle"))
                OrbHwDocTool.NewDocCustomProperty("orbDocShortTittle", Office.MsoDocProperties.msoPropertyTypeString, "Document Short Tittle");

            if (!OrbHwDocTool.CustomPropertyExist("orbDocContract"))
                OrbHwDocTool.NewDocCustomProperty("orbDocContract", Office.MsoDocProperties.msoPropertyTypeString, "Contract");

            if (!OrbHwDocTool.CustomPropertyExist("orbDocContractTittle"))
                OrbHwDocTool.NewDocCustomProperty("orbDocContractTittle", Office.MsoDocProperties.msoPropertyTypeString, "Contract tittle");

            if (!OrbHwDocTool.CustomPropertyExist("orbDocProgram"))
                OrbHwDocTool.NewDocCustomProperty("orbDocProgram", Office.MsoDocProperties.msoPropertyTypeString, "Program");

            if (!OrbHwDocTool.CustomPropertyExist("orbDocProject"))
                OrbHwDocTool.NewDocCustomProperty("orbDocProject", Office.MsoDocProperties.msoPropertyTypeString, "Project");

            if (!OrbHwDocTool.CustomPropertyExist("orbDocDrlNum"))
                OrbHwDocTool.NewDocCustomProperty("orbDocDrlNum", Office.MsoDocProperties.msoPropertyTypeNumber, 0);

            #endregion

            #region PROPIEDADES DE LA VERSIÓN INICIAL (1.0)

            // En el caso de un documento nuevo hay que generar una primera versión de partida
            // Creamos la propiedad contadora de versiones y la inicializamos a cero
            if (!OrbHwDocTool.CustomPropertyExist("orbDocIssueNum"))
            {
                lastIssueNum = 0;
                OrbHwDocTool.NewDocCustomProperty("orbDocIssueNum", Office.MsoDocProperties.msoPropertyTypeNumber, lastIssueNum);

                /*** EDICION (MAJOR) ***/
                if (!OrbHwDocTool.CustomPropertyExist("orbDocIssueMajor_0"))
                    OrbHwDocTool.NewDocCustomProperty("orbDocIssueMajor_0", Office.MsoDocProperties.msoPropertyTypeNumber, 1);

                /*** REVISIÓN (MINOR) ***/
                if (!OrbHwDocTool.CustomPropertyExist("orbDocIssueMinor_0"))
                    OrbHwDocTool.NewDocCustomProperty("orbDocIssueMinor_0", Office.MsoDocProperties.msoPropertyTypeNumber, 0);

                /*** FECHA ***/
                if (!OrbHwDocTool.CustomPropertyExist("orbDocIssueDate_0"))
                    OrbHwDocTool.NewDocCustomProperty("orbDocIssueDate_0", Office.MsoDocProperties.msoPropertyTypeDate, 
                        DateTime.Now);

                /*** MOTIVO ***/
                    OrbHwDocTool.NewDocCustomProperty("orbDocIssuerReason_0", Office.MsoDocProperties.msoPropertyTypeString,
                        "New document.");
            }
            #endregion

            #region PROPIEDADES DE LA VERSION ACTUAL DEL DOCUMENTO

            /*** VERSION (MAJOR.MINOR) ***/
            if (!OrbHwDocTool.CustomPropertyExist("orbDocIssue"))
                OrbHwDocTool.NewDocCustomProperty("orbDocIssue", Office.MsoDocProperties.msoPropertyTypeString, "1.0");

            /*** FECHA ***/
            if (!OrbHwDocTool.CustomPropertyExist("orbDocIssueDate"))
                OrbHwDocTool.NewDocCustomProperty("orbDocIssueDate", Office.MsoDocProperties.msoPropertyTypeDate,
                    DateTime.Now);

            #endregion

            OrbHwDocTool.UpdateAllDocFields();  // Actualizamos todos los campos del documento
        }

        public static void NewDocCustomProperty(string prop, Office.MsoDocProperties type, object content)
        {
            Office.DocumentProperties toolDocCustomProps =
                Globals.ThisDocument.CustomDocumentProperties as Office.DocumentProperties;

            if (!CustomPropertyExist(prop))
                toolDocCustomProps.Add(prop, false, type, content);
        }

        public static Boolean CustomPropertyExist(string propName)
        {
            Office.DocumentProperties toolDocCustomProps =
                Globals.ThisDocument.CustomDocumentProperties as Office.DocumentProperties;

            try
            {
                Office.DocumentProperty temp = toolDocCustomProps[propName];
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void UpdateAllDocFields()
        {
            foreach (Word.Range range in Globals.ThisDocument.StoryRanges)
            {
                Word.Range r = range;

                while (r != null)
                {
                    r.Fields.Update();
                    r = r.NextStoryRange;       // return null at the end.
                }
            }
        }
    }
}