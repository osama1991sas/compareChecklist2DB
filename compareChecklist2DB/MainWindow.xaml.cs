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
using System.Collections.Generic;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Net.NetworkInformation;
using System.Diagnostics;


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
            {
            //var dialog = new Microsoft.Win32.OpenFileDialog
            //{
            //    Title = "Select a Folder",
            //    CheckFileExists = false,
            //    CheckPathExists = true,
            //    FileName = "Select Folder",
            //};

            //if (dialog.ShowDialog() == true)
            //{
            //    string selectedFolder = System.IO.Path.GetDirectoryName(dialog.FileName); // Correct usage
            //    SelectedFolderText.Text = selectedFolder;
            //}
            }
            var dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Title = "Select a folder"
            };

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string selectedFolder = dialog.FileName;
                SelectedFolderText.Text = selectedFolder;
                // Do something with selectedPath
            }
        }
        string outputDir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CleanCSV");

        private void CompareChecklist2DB_Click(object sender, RoutedEventArgs e)
        {
            try {
            File.WriteAllText("Missing Devices.txt", string.Empty);}
            catch { }
            Output.Text = "Missing Devices: \n\n";
            // Register the encoding provider to handle Windows-1252 encoding
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string folderPath = SelectedFolderText.Text; // Update this to your folder path
            List<string> csvfiles = new List<string>();
            List<string> devices = new List<string>();
            List<string> chklistDevices = new List<string>();
            // Getting all devices from csv file
            try
            {
                using (StreamReader sr = new StreamReader(SelectedFileText.Text))
                {
                    sr.ReadLine(); // Line 1
                    sr.ReadLine(); // Line 2
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        string[] columns = line.Split(','); // Split by comma
                        if (columns.Length > 1) // Ensure there is a second column
                            devices.Add(columns[2].Substring(columns[2].IndexOf('L')));
                    }
                }

                // Cleaning the checklist files
                foreach (string file in Directory.GetFiles(folderPath, "*.xls"))
                {
                    DataTable dataTable = Helper.ReadExcelFile(file);
                    DataTable cleanedTable = Helper.RemoveFirstThreeRows(dataTable);
                    if (!Directory.Exists(outputDir))
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
                    string deviceType = csvNameSplit[csvNameSplit.Length - 1][0] + "";
                    using (StreamReader sr = new StreamReader(file))
                    {
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            string[] columns = line.Split(','); // Split by comma
                            if (columns.Length > 1) // Ensure there is a second column
                                if (columns[1].Contains("True"))
                                    chklistDevices.Add("L" + loopNum + deviceType + columns[0].PadLeft(3, '0'));
                        }
                    }
                }

                foreach (string chklistdevice in chklistDevices)
                    if (!devices.Contains(chklistdevice) && !Regex.IsMatch(chklistdevice, @"L01M15[5-9]"))
                        using (StreamWriter writer = new StreamWriter("Missing Devices.txt", true)) // 'true' enables appending
                        {
                            writer.WriteLine(chklistdevice);
                            Output.AppendText(chklistdevice + "\n");
                            
                        }
                //rewrite first line to include the number of missing devices
                string fullText = Output.Text;
                int numOfLines = Output.LineCount - 3;
                // Get the index of the first line's start and end
                int start = Output.GetCharacterIndexFromLineIndex(0);
                int length = Output.GetLineLength(0);

                // Replace the first line with your new content
                string newFirstLine = "Missing Devices: " + numOfLines + "\n\n";
                string restOfText = fullText.Substring(start + length);

                Output.Text = newFirstLine + restOfText;
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        private void CheckDBIntegrity_Click(object sender, RoutedEventArgs e)
        {

            List<string> deviceNameIssue = new List<string>();
            List<string> zoneNameIssue = new List<string>();
            List<string> deviceNameAndAddress = new List<string>();
            List<string> roomNameLong = new List<string>();
            List<string> devices = new List<string>();

            try{ using (StreamReader sr = new StreamReader(SelectedFileText.Text))
            {
                sr.ReadLine(); // Line 1
                sr.ReadLine(); // Line 2
                string line;
                while ((line = sr.ReadLine()) != null)
                {

                    string[] columns = line.Split(','); // Split by comma
                    if (line.StartsWith("ADDRESS") && columns.Length > 6)
                        Output.AppendText("Room Name & Number Separated\n\n");
                    else if (columns.Length > 1) // Ensure there is a second column
                    {
                        if (!Regex.IsMatch(columns[2], @"N\d{1,3}L0\d[DM]\d{3}$"))
                            deviceNameIssue.Add(line);
                        if (!Regex.IsMatch(columns[1], @"ZONE-(?:[A-Z]\d{2,3}|\d{2,3})\b"))
                            zoneNameIssue.Add(line);
                        var addressParts = Regex.Matches(columns[2], @"[A-Z]\d+").Cast<Match>().Select(m => m.Value).ToArray();
                        var fullParts = columns[3].Split('-');
                        if (!addressParts[addressParts.Length - 1].Contains(fullParts[fullParts.Length - 1]) ||
                            !(fullParts[fullParts.Length - 2].Contains(addressParts[addressParts.Length - 1][0]) || addressParts[addressParts.Length - 1][0] == 'D') ||
                            !fullParts[fullParts.Length - 3].Substring(2).Contains(addressParts[addressParts.Length - 2].Substring(1)))
                            deviceNameAndAddress.Add(line);
                        if (columns[4].Length > 35) {
                            roomNameLong.Add(line);
                        }
                    }
                    devices.Add(columns[2].Substring(4));
                }


            }
            var duplicates = devices.GroupBy(s => s).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
            if (deviceNameIssue.Count > 0 || zoneNameIssue.Count > 0 || deviceNameAndAddress.Count > 0 || roomNameLong.Count > 0 || duplicates.Count > 0)
            {
                Output.AppendText("Device with Address Issue: \n\n");
                foreach (var issue in deviceNameIssue) { Output.AppendText(issue + "\n"); }
                Output.AppendText("\n\nDevice with ZONE Issue: \n\n");
                foreach (var issue in zoneNameIssue) { Output.AppendText(issue + "\n"); }
                Output.AppendText("\n\nDevice with Address Miss Match: \n\n");
                foreach (var issue in deviceNameAndAddress) { Output.AppendText(issue + "\n"); }
                Output.AppendText("\n\nRoom Name Too Long: \n\n");
                foreach (var issue in roomNameLong) { Output.AppendText(issue + "\n"); }
                Output.AppendText("\n\nDuplicated Devices: \n\n");
                foreach (var issue in duplicates) { Output.AppendText(issue + "\n"); }
            }
            else
                MessageBox.Show("No Issues detected in DB");
        }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void ClearOutput(object sender, RoutedEventArgs e)
        {
            MessageBoxResult mbr = System.Windows.MessageBox.Show("Do you also want clear Missing Devices.txt and CleanCSV folder?","Clearing Data"
                ,MessageBoxButton.YesNoCancel);

            if (mbr == MessageBoxResult.Yes)
            {
                try
                {
                    File.WriteAllText("Missing Devices.txt", string.Empty);
                }
                catch { }
                // Delete all files
                foreach (string file in Directory.GetFiles(outputDir))
                {
                    File.Delete(file);
                }

                // Delete all subdirectories and their contents
                foreach (string dir in Directory.GetDirectories(outputDir))
                {
                    Directory.Delete(dir, true); // true means delete subdirectories and files recursively
                }
                Output.Text = string.Empty;
            }
            else if (mbr == MessageBoxResult.No)
                Output.Text = string.Empty;
            else { }
        }
        private void OpenFolder(object sender, RoutedEventArgs e)
        {
            Process.Start("explorer.exe", AppDomain.CurrentDomain.BaseDirectory);
        }
    }

}