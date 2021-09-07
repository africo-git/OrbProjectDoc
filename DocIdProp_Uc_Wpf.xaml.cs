using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace OrbProjectDoc
{
    /// <summary>
    /// Lógica de interacción para DocIdProp_Uc_Wpf.xaml
    /// </summary>
    public partial class DocIdProp_Uc_Wpf : UserControl
    {
        public DocIdProp_Uc_Wpf()
        {
            InitializeComponent();
        }

        private void CmdUpdateFormFields_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisDocument.myDocIdProp_Uc.UpdateFrom();
        }

        private void CmdSaveCustomDocProperties_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisDocument.myDocIdProp_Uc.SaveChange();
        }
    }
}
