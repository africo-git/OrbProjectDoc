using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Windows.Forms;


namespace OrbProjectDoc
{
    public partial class RibbonTemplate
    {
        private void RibbonTemplate_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void TglBtoDocIdProp_Click(object sender, RibbonControlEventArgs e)
        {
            //// Obtenemos el índice del control del ActionsPane que queremos mostrar.
            //int myDocIdProp_Uc_index = Globals.ThisDocument.ActionsPane.Controls.GetChildIndex(Globals.ThisDocument.myDocIdProp_Uc);

            //if (TglBtoDocIdProp.Checked)
            //{
            //    // Mostramos el control desesado del ActionsPane (cargado al inicio del documento)
            //    Globals.ThisDocument.ActionsPane.Controls[myDocIdProp_Uc_index].Visible = true;

            //    // Mostramos el Document Actions Task Pane
            //    Globals.ThisDocument.Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneDocumentActions].Visible = true;

            //    // Gestionamos la visibilidd de los controles del Ribbon que no sean compatibles
            //    TglBtoCtrlVer.Enabled = false;
            //}
            //else
            //{
            //    // Ocultamos el control desesado del ActionsPane (cargado al inicio del documento)
            //    Globals.ThisDocument.ActionsPane.Controls[myDocIdProp_Uc_index].Visible = false;

            //    // Ocultamos el Document Actions Task Pane
            //    Globals.ThisDocument.Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneDocumentActions].Visible = false;

            //    // Gestionamos la visibilidd de los controles del Ribbon que no sean compatibles
            //    TglBtoCtrlVer.Enabled = true;
            //}
        }

        private void TglBtoCtrlVer_Click(object sender, RibbonControlEventArgs e)
        {
            //// Obtenemos el índice del control del ActionsPane que queremos mostrar.
            //int myDocVerCtrl_Uc_index = Globals.ThisDocument.ActionsPane.Controls.GetChildIndex(Globals.ThisDocument.myDocVerCtrl_Uc);

            //if (TglBtoCtrlVer.Checked)
            //{
            //    // Mostramos el control desesado del ActionsPane (cargado al inicio del documento)
            //    Globals.ThisDocument.ActionsPane.Controls[myDocVerCtrl_Uc_index].Visible = true;

            //    // Mostramos el Document Actions Task Pane
            //    Globals.ThisDocument.Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneDocumentActions].Visible = true;

            //    // Gestionamos la visibilidd de los controles del Ribbon que no sean compatibles
            //    TglBtoDocIdProp.Enabled = false;    // Desactivado otros botones incompatibles
            //}
            //else
            //{
            //    // Ocultamos el control desesado del ActionsPane (cargado al inicio del documento)
            //    Globals.ThisDocument.ActionsPane.Controls[myDocVerCtrl_Uc_index].Visible = false;

            //    // Ocultamos el Document Actions Task Pane
            //    Globals.ThisDocument.Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneDocumentActions].Visible = false;

            //    // Gestionamos la visibilidd de los controles del Ribbon que no sean compatibles
            //    TglBtoDocIdProp.Enabled = true;     // Activa los botones que eran incompatibles
            //}
        }
    }
}
