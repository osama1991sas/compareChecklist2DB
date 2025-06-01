using Microsoft.Win32;
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
using System.IO;
using System.Data;
using System.Text.RegularExpressions;
using ExcelDataReader;
using compareChecklist2DB.Classes;

namespace compareChecklist2DB
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
        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select a File";
            openFileDialog.Filter = "All Files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true) // If user selects a file
            {
                SelectedFileText.Text = openFileDialog.FileName;
            }
        }
        private void SelectFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select a Folder",
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "Select Folder"
            };

            if (dialog.ShowDialog() == true)
            {
                string selectedFolder = System.IO.Path.GetDirectoryName(dialog.FileName); // Correct usage
                SelectedFolderText.Text = selectedFolder;
            }
        }

        private void CompareChecklist2DB_Click(object sender, RoutedEventArgs e)
        {
            Output.Text = "Missing Devices: \n\n";
            // Register the encoding provider to handle Windows-1252 encoding
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string folderPath = SelectedFolderText.Text; // Update this to your folder path
            string outputDir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CleanCSV");
            List<string> csvfiles = new List<string>();
            List<string> devices = new List<string>();
            List<string> chklistDevices = new List<string>();
            // Getting all devices from csv file
            using (StreamReader sr = new StreamReader(SelectedFileText.Text))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    string[] columns = line.Split(','); // Split by comma
                    if (columns.Length > 1) // Ensure there is a second column
                        devices.Add(columns[2].Substring(4));
                }
                devices.RemoveAt(0);
            }

            // Cleaning the checklist files
            foreach (string file in Directory.GetFiles(folderPath, "*.xls"))
            {
                DataTable dataTable = Helper.ReadExcelFile(file);
                DataTable cleanedTable = Helper.RemoveFirstThreeRows(dataTable);
                if(!Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);
                string csvFilePath = System.IO.Path.Combine(outputDir, System.IO.Path.GetFileNameWithoutExtension(file) + ".csv");
                Helper.SaveAsCsv(cleanedTable, csvFilePath);
                csvfiles.Add(csvFilePath);
            }
            // checking missing devices
            foreach (string file in csvfiles)
            {
                string csvNames = file.Substring(file.IndexOf("CleanCSV") + 9);
                string loopNum = Regex.Match(csvNames, @"\b(?:L|LOOP)\s*(\d+)\b").Groups[1].Value.PadLeft(2, '0');
                string[] csvNameSplit = csvNames.Split(" ");
                string deviceType = csvNameSplit[csvNameSplit.Length - 1][0]+"";
                using (StreamReader sr = new StreamReader(file))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        string[] columns = line.Split(','); // Split by comma
                        if (columns.Length > 1) // Ensure there is a second column
                            if (columns[1].Contains("True"))
                                chklistDevices.Add("L" + loopNum+ deviceType + columns[0].PadLeft(3, '0'));
                    }
                }
            }
            foreach (string chklistdevice in chklistDevices)
                if (!devices.Contains(chklistdevice))
                    using (StreamWriter writer = new StreamWriter("Missing Devices.txt", true)) // 'true' enables appending
                    {
                        writer.WriteLine(chklistdevice);
                        Output.AppendText(chklistdevice+"\n");
                    }
        }

        private void ClearOutput(object sender, RoutedEventArgs e)
        {
            Output.Text = string.Empty;
        }

    }

}