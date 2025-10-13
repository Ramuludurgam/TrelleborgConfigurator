using Microsoft.VisualBasic;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Configurator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnGen_Click(object sender, RoutedEventArgs e)
        {

            try
            {

                SldWorks swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
                // Try to get running instance

                if (swApp != null)
                {
                    swApp.Visible = true;
                    //swApp.SendMsgToUser2("Got the handle.", (int)swMessageBoxIcon_e.swMbInformation, (int)swMessageBoxBtn_e.swMbOk);

                    if ( swApp.GetDocumentCount() > 0)
                    {
                        MessageBox.Show("Close all the opened files.", "Trelleborg", button: MessageBoxButton.OK, icon: MessageBoxImage.Warning);
                        return;
                    }


                    string destPath = this.GetFolderPath();
                   
                    if (destPath is null)
                    {
                        MessageBox.Show("Path to store the 3D files is required", "Trelleborg",button: MessageBoxButton.OK,icon: MessageBoxImage.Warning);
                        return;
                    }

                    object[] files =  Directory.GetFiles(destPath);

                    if (files.Length > 0)
                    {
                        MessageBoxResult result = MessageBox.Show("There are files under this folder that may get overriden do you want to continue", "Trelleborg", button: MessageBoxButton.YesNo, icon: MessageBoxImage.Warning);

                        if (result == MessageBoxResult.No) 
                            { return; }
                    }

                    TrelleborgAssembly trelleborgAssembly = new TrelleborgAssembly(swApp, destPath);
                    
                    trelleborgAssembly.CreateRequiredPartFiles(); //First create the required files for this assembly
                    trelleborgAssembly.CreateFullAssembly(); // generate the assembly after generating all the part files..

                    trelleborgAssembly.CreateDrawingFile();

                    MessageBox.Show("All the required files are created successfully.", "Trelleborg", button: MessageBoxButton.OK, icon: MessageBoxImage.None);
                }
            }
            catch (COMException)
            {
                MessageBox.Show("SOLIDWORKS is not running.");
            }

        }
        private string GetFolderPath()
        {
            OpenFolderDialog folderDialog = new OpenFolderDialog();

            folderDialog.Multiselect = false;
            folderDialog.Title = "Select a folder";

            bool? result = folderDialog.ShowDialog();

            if (result == true)
            {
                return folderDialog.FolderName;
            }

            return null;

        }
    }
}
