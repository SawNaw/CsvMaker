using CsvMaker;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Windows.Documents;
using System.Windows.Input;
using System.Text.RegularExpressions;
using System.Windows.Media;
using Gat.Controls;


namespace CsvMaker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<string> inputFiles { get; set; } // user-selected file(s) for processing
        private string _outputDirectory; // user-selected path for converted CSV file
        public string OutputDirectory
        {
            get { return _outputDirectory; }
            set
            {
                _outputDirectory = value;
                if (value != null)
                    WriteToStatusDisplay($"{Environment.NewLine}Save directory set to: {_outputDirectory}", 
                        "Files will be created under this directory.", Brushes.Brown);
            }
        }

        public MainWindow()
        {
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;  // Center the window on startup

            // Set tooltip properties for this window. This can't be done in XAML because it gets overriden somewhere.
            ToolTipService.ShowDurationProperty.OverrideMetadata(typeof(DependencyObject), new FrameworkPropertyMetadata(Int32.MaxValue));
            ToolTipService.InitialShowDelayProperty.OverrideMetadata(typeof(DependencyObject), new FrameworkPropertyMetadata(50));
            ToolTipService.BetweenShowDelayProperty.OverrideMetadata(typeof(DependencyObject), new FrameworkPropertyMetadata(50));
            InitializeComponent();

            // Initialize member variables
            inputFiles = new List<string>();
            OutputDirectory = null;

            // Display ready status in status box
            WriteToStatusDisplay("Ready.", Brushes.Brown);

            // Display tip in field lengths text box
            tbFieldLengths.Foreground = Brushes.Gray;
            tbFieldLengths.FontStyle = FontStyles.Italic;
            tbFieldLengths.Text = "Enter field lengths separated by commas (e.g. 15,31,24)";
        }

        /// <summary>
        /// Prepares to convert a delimited file to CSV and notifies the user of progress through the status display.
        /// </summary>
        /// <param name="file">The full path of the delimited file to be converted.</param>
        /// <param name="delimiter">The delimiting character used in the delimited file.</param>
        /// <returns>True if no errors were encountered, false otherwise.</returns>
        private bool BeginDelimitedToCsvConversion(List<string> files, char delimiter, string qualifier="")
        {
            bool noErrors = true;
            string currentTime = DateTime.Now.ToString();
            foreach (string file in files)
            {
                // Print current activity in the status display
                string msg = $"{Environment.NewLine}Converting {Path.GetFileName(file)} ...";
                string tooltip = $"Started {DateTime.Now.ToString()}";
                Dispatcher.Invoke( () => WriteToStatusDisplay(msg, tooltip, Brushes.DarkOrange), DispatcherPriority.Background );

                // Construct output file path and output a CSV file to the path
                string fullOutputPath = $"{_outputDirectory}\\{Path.GetFileNameWithoutExtension(file)}.csv";

                var error = CsvConverter.WriteCsvRowsToFile(delimiter, file, fullOutputPath, qualifier);
                if (error != null) // errors encountered during operation
                {
                    noErrors = false;
                    File.Delete(fullOutputPath);  // delete the incomplete CSV file
                    msg = " Error!";
                    tooltip = $"Error on line {error.LineNumber}: {error.Message}";
                    Dispatcher.Invoke(() => WriteErrorToStatusDisplay(msg, tooltip), DispatcherPriority.Background);
                }
                else // no errors detected
                {
                    msg = " Done!";
                    tooltip = $"This file was successfully processed at {DateTime.Now.ToString("T")}";
                    Dispatcher.Invoke(() => WriteToStatusDisplay(msg, tooltip, Brushes.Green), DispatcherPriority.Background);
                }
            }
            return noErrors;
        }

        /// <summary>
        /// Prepares to convert a fixed length file to CSV and notifies the user of progress through the status display.
        /// </summary>
        /// <param name="file">The full path of the file to be converted.</param>
        /// <returns>True if no errors were encountered. False otherwise.</returns>
        private async Task<bool> BeginFixedLengthToCsvConversionAsync(List<string> files, string qualifier="")
        {
            bool noErrors = false;
            // Get the field count and field lengths from user input
            int numFields = tbFieldLengths.Text.Split(',').Length;  // Number of field lengths entered in tbFieldLengths
            string[] fieldWidthsAsStrings = tbFieldLengths.Text.Split(',');
            int[] fieldWidths = Array.ConvertAll(fieldWidthsAsStrings, int.Parse);
            string currentTime = DateTime.Now.ToString();

            foreach (string file in files)
            {
                // Print current activity in the status display
                string msg = $"{Environment.NewLine}Converting {Path.GetFileName(file)} ...";
                string tooltip = $"Started {DateTime.Now.ToString()}";
                WriteToStatusDisplay(msg, tooltip, Brushes.DarkOrange);

                // Construct output file path 
                string fullOutputPath = $"{OutputDirectory}\\{Path.GetFileNameWithoutExtension(file)}.csv";
                
                // Start writing CSV rows to the file
                long resultCode = await Task.Run(() => CsvConverter.WriteCsvRowsToFile(fieldWidths, file, fullOutputPath, qualifier));
                if (resultCode != 0)  // Error converting file to CSV. At present, only a MalformedLineException will cause this.
                {
                    noErrors = true;
                    string errorMsg = $"{Environment.NewLine}   Error converting {file}: {Environment.NewLine}   Line {resultCode} cannot be parsed with the given field lengths.";
                    string tooltipText = "Check that the field lengths you specified are correct for this file. Note that all selected files must have the same field lengths.";
                    WriteErrorToStatusDisplay(errorMsg, tooltipText);
                }
                else
                {
                    WriteToStatusDisplay(" Done!", "This file was successfully processed.", Brushes.Green);
                }
            }
            return noErrors;
        }

        /// <summary>
        /// TO DO: IMPROVE STATUS OUTPUT!!!
        /// Prepares to convert an Excel file to CSV and notifies the user of progress through the status display.
        /// </summary>
        /// <param name="files">The full path of the file to be converted.</param>
        /// <param name="excelFileType">The type of excel file type. Value should be either ".xls" or ".xlsx".</param>
        /// <returns>True if no errors were encountered, false otherwise.</returns>
        private async Task<bool> BeginExcelToCsvConversionAsync(List<string> files, string excelFileType, string qualifier="")
        {
            bool noErrors = true;  // false if any exceptions were caught
            foreach (string file in files)
            {
                // Print current activity in status display
                string msg = $"{Environment.NewLine}Converting {Path.GetFileName(file)} ...";
                string tooltip = "Started " + DateTime.Now.ToString();
                WriteToStatusDisplay(msg, tooltip, Brushes.DarkOrange);

                // Form connection string for the Excel file
                string connectionString = null;
                switch (excelFileType.ToLower())
                {
                    case ".xls":
                        connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.16.0; 
                                                            Data Source={0}; OLE DB Services=-1;
                                                            Extended Properties='Excel 8.0; HDR=Yes;'"
                                                            , file); 
                                                            // HDR=Yes indicates that the first row contains column names, not data
                        break;
                    case ".xlsx":
                        connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.16.0;
                                                            Data Source={0}; OLE DB Services=-1;
                                                            Extended Properties='Excel 12.0 Xml; HDR=YES; IMEX=1;'"
                                                            // IMEX=1 treats all data as text
                                                            , file);
                        break;
                    default:
                        string errorMsg = "Fatal error: An Excel radio button was expected from parameter \"excelFileType\" in ConvertExcelFilesToCsv() but " +
                                        "not received. Please contact the developer.";
                        string caption = "Fatal error in ConvertExcelFilesToCsv()";
                        MessageBox.Show(errorMsg, caption);
                        break;
                }

                // Create new OleDB connection 
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        // Open the OleDB connection to the Excel sheet and execute query
                        await Task.Run( () => connection.Open() );
                        string queryString = "SELECT * FROM [Sheet1$]";
                        OleDbCommand command = new OleDbCommand(queryString, connection);

                        string currentTime = DateTime.Now.ToString();

                        // Construct the output file path and begin writing CSV rows to that file
                        string fullOutputPath = OutputDirectory + @"\" + Path.GetFileNameWithoutExtension(file) + ".csv";
                        CsvConverter.WriteCsvRowsToFile(command, fullOutputPath, qualifier);
                    }
                    catch (OleDbException e)
                    {
                        noErrors = false;
                        string errorMsg = $"{Environment.NewLine}   Error: " + e.Message;
                        string tip = "There is a problem with this Excel file. Please refer to the status display.";
                        WriteErrorToStatusDisplay(errorMsg, tip);
                    }
                    WriteToStatusDisplay(" Done!", "This file was successfully processed.", Brushes.Green);
                }
            }
            return noErrors;
        }

        /// <summary>
        /// Writes a message to the status message display with a red foreground and light yellow background, and with an associated tooltip.
        /// </summary>
        /// <param name="msgText">The text to write to the message display.</param>
        /// <param name="tooltipText">The tooltip to display when user mouses over the message.</param>
        private void WriteErrorToStatusDisplay(string msgText, string tooltipText)
        {
            var richTextMsg = new Bold(new Run(msgText));
            richTextMsg.Foreground = Brushes.Red;
            richTextMsg.ToolTip = tooltipText;
            tblStatus.Inlines.Add(richTextMsg);
        }

        /// <summary>
        /// Writes a message to the status message display with the specified tooltip, and foreground and background colours.
        /// </summary>
        /// <param name="msgText">The text to write to the message display.</param>
        /// <param name="tooltipText">The tooltip to display when user mouses over the message.</param>
        /// <param name="foregroundColor">The foreground colour of the message.</param>
        /// <param name="backgroundColor">The background colour of the message.</param>
        private void WriteToStatusDisplay(string msgText, string tooltipText, Brush foregroundColor, Brush backgroundColor)
        {
            var richTextMsg = new Bold(new Run(msgText));
            richTextMsg.Foreground = foregroundColor;
            richTextMsg.Background = backgroundColor;
            richTextMsg.ToolTip = tooltipText;
            tblStatus.Inlines.Add(richTextMsg);
        }
        
        /// <summary>
        /// Writes a message to the status message display with the specified tooltip, and foreground colour.
        /// </summary>
        /// <param name="msgText">The text to write to the message display.</param>
        /// <param name="tooltipText">The tooltip to display when user mouses over the message.</param>
        /// <param name="foregroundColor">The foreground colour of the message.</param>
        private void WriteToStatusDisplay(string msgText, string tooltipText, Brush foregroundColor)
        {
            var richTextMsg = new Bold(new Run(msgText));
            richTextMsg.Foreground = foregroundColor;
            richTextMsg.ToolTip = tooltipText;
            tblStatus.Inlines.Add(richTextMsg);
        }

        /// <summary>
        /// Writes a message to the status message display with the specified foreground colour.
        /// </summary>
        /// <param name="msgText">The text to write to the message display.</param>
        /// <param name="foregroundColor">The foreground colour of the message.</param>
        private void WriteToStatusDisplay(string msgText, Brush foregroundColor)
        {
            var richTextMsg = new Bold(new Run(msgText));
            richTextMsg.Foreground = foregroundColor;
            tblStatus.Inlines.Add(richTextMsg);
        }

        /// <summary>
        /// Checks the appropriate radio box based on the selected files' extension
        /// </summary>
        /// <param name="file">The file extension of the selected files</param>
        private void CheckRadioBoxBasedOnFileExtension(string file)
        {
            string fileExtension = Path.GetExtension(file).ToLower();
            switch (Path.GetExtension(file).ToLower())
            {
                case ".fix":
                        rbnFixed.IsChecked = true;
                        break;
                case ".xls":
                        rbnExcelXls.IsChecked = true;
                        break;
                case ".xlsx":
                        rbnExcelXlsx.IsChecked = true;
                        break;
                case "csv":
                        rbnDelimited.IsChecked = true;
                        break;
                case ".txt":
                        // Idea: Possibly employ heuristic methods to detect what kind of text file it is and process accordingly?
                        break; 
                default:
                    break;
            }
        }

        /// <summary>
        /// Reads a layout file (e.g. K2 Direct Export Layout File) and extracts the field lengths.
        /// </summary>
        /// <param name="path">The full path of the layout file.</param>
        private static List<string> GetFieldLengthsFromLayoutFile(string path)
        {
            List<string> fieldLengths = new List<string>();
            foreach (string s in File.ReadLines(path).Skip(8))
            {
                string[] lines = s.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                fieldLengths.Add(lines[1]);
            }

            return fieldLengths;
        }

        #region User Interface Event Handlers

        private void SaveCmdExecuted(object sender, RoutedEventArgs e)
        {
            // Set up a folder browser
            var folderbrowserDialog = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();
            folderbrowserDialog.IsFolderPicker = true;
            // DialogResult result = dlg.ShowDialog();
            if (folderbrowserDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                OutputDirectory = folderbrowserDialog.FileName;
            }
            lblOutputFilePath.Content = OutputDirectory;
        }

        // Executes when user clicks the Load File menu item or button.
        private void OpenCmdExecuted(object sender, RoutedEventArgs e)
        {
            // Set up an OpenFileDialog object
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true; // Note: batch processing is not currently available for files of mixed types
            dlg.Filter = "All supported types (*.txt,*.s01,*.xls,*.xlsx,*.csv)|*.TXT;*.S01;*.XLS;*.XLSX;*.CSV|"
                            + "Text files (*.txt)|*.TXT|"
                            + "CSV files (*.csv)|*.TXT|"
                            + "s01 files (*.s01)|*.S01|"
                            + "Excel 98-2003 files (*.xls)|*.XLS|"
                            + "Excel 2007 and Later Files (*.xlsx)|*.XLSX|"
                            + "All Types|*.*";

            // Open the OpenFileDialog box
            bool? userClickedOk = dlg.ShowDialog();
            if (userClickedOk == true)
            {
                // Check if the selected files contain mixed types. Mixed types are currently not supported.
                if (!AllFilesHaveSameExtension(dlg.FileNames))
                {
                    string msg = "Batch processing of different file types is not supported. Please select files that have the same extension and type.";
                    string caption = "Mixed file types not supported.";
                    MessageBox.Show(this, msg, caption, MessageBoxButton.OK, MessageBoxImage.Hand);
                }
                else // file selection is valid, so begin processing
                {
                    // Add the user-selected  files to the processing list
                    this.inputFiles = dlg.FileNames.ToList<string>();

                    // Set the default output file directory to a timestamped CsvMaker directory under the current directory
                    string currentDirectoryName = Path.GetDirectoryName(dlg.FileName);
                    OutputDirectory = currentDirectoryName;

                    // Update the UI to reflect the loaded files.
                    lblSourceFilePath.Content = $"Files loaded: {inputFiles.Count}";
                    WriteToStatusDisplay($"{Environment.NewLine}Files loaded: {inputFiles.Count}", Brushes.Brown);
                    lblOutputFilePath.Content = OutputDirectory;

                    // Get the extension of all selected files. 
                    // Note: All selected files must have the same extension, for now.
                    string firstFileInSelection = inputFiles.ElementAt<string>(0);
                    CheckRadioBoxBasedOnFileExtension(firstFileInSelection);
                }
            }
            if (ValidateInputSuccess())
                EnableCreateCsvButton();
            else
                DisableCreateCsvButton();
        }
        
        /// <remarks>The only methods that should be async void are event handlers.</remarks>
        private async void CreateCsvClicked(object sender, RoutedEventArgs e)
        {
            string qualifier = null;
            if (cbEncloseWithQualifiers.IsChecked == true) // user specified that each field in CSV file should be enclosed with qualifiers
                qualifier = tbQualifier.Text;  // get the qualifier character
            
            // Validate input based on which radio button is selected. If validation is successful, perform conversion.
            if (rbnDelimited.IsChecked == true) // User specified that file is delimited
            {
                if (ValidateInputSuccess()) // Check whether parameters of the delimited file were correctly specified
                {
                    string msg = $"{Environment.NewLine}{Environment.NewLine}Converting delimited data to CSV...{Environment.NewLine}" +
                                    $"==========================================={Environment.NewLine}";
                    string tooltipText = "Job started " + DateTime.Now.ToString();
                    Dispatcher.Invoke( () => WriteToStatusDisplay(msg, tooltipText, Brushes.Black), DispatcherPriority.Background );
                    
                    // Get the delimiter from the combo box.
                    char delimiter = '\0';
                    if (cbiPipe.IsSelected) delimiter = '|';
                    else if (cbiSemicolon.IsSelected) delimiter = ';';
                    else if (cbiTab.IsSelected) delimiter = '\t';
                    else if (cbiOther.IsSelected) delimiter = tbDelimiter.Text[0];

                    bool x = await Task.Run( () => BeginDelimitedToCsvConversion(this.inputFiles, delimiter, qualifier));
                }

            }
            else if (rbnFixed.IsChecked == true) // User specified that file is fixed length
            {
                if (ValidateInputSuccess()) // Check whether parameters of the fixed length file were correctly specified
                {
                    string msg = $"{Environment.NewLine}{Environment.NewLine}Converting fixed length data to CSV...{Environment.NewLine}" +
                                    $"============================================{Environment.NewLine}";
                    string tooltipText = "Job started " + DateTime.Now.ToString();
                    WriteToStatusDisplay(msg, tooltipText, Brushes.Black);
                    bool x = await BeginFixedLengthToCsvConversionAsync(this.inputFiles, qualifier);
                }
            }
            else if (rbnExcelXls.IsChecked == true) // User specified that file is Excel (.xls)
            {
                if (ValidateInputSuccess()) // Check whether parameters of the delimited file were correctly specified
                {
                    string msg = $"{Environment.NewLine}{Environment.NewLine}Converting Excel (.xls) to CSV...{Environment.NewLine}" +
                                    $"============================================{Environment.NewLine}";
                    string tooltipText = "Job started " + DateTime.Now.ToString();
                    WriteToStatusDisplay(msg, tooltipText, Brushes.Black);

                    bool x = await BeginExcelToCsvConversionAsync(this.inputFiles, ".xls", qualifier);
                }
            }
            else if (rbnExcelXlsx.IsChecked == true) // User specified that file is Excel (.xlsx)
            {
                if (ValidateInputSuccess()) // Check whether parameters of the delimited file were correctly specified
                {
                    string msg = $"{Environment.NewLine}{Environment.NewLine}Converting Excel (.xlsx) to CSV...{Environment.NewLine}" +
                                       $"==========================================={Environment.NewLine}";
                    string tooltipText = "Job started " + DateTime.Now.ToString();
                    WriteToStatusDisplay(msg, tooltipText, Brushes.Black);
                    bool x = await BeginExcelToCsvConversionAsync(this.inputFiles, ".xlsx", qualifier);
                }
            }

            // This code should be reached only after processing has been successfully completed.
            if (ValidateInputSuccess())
            {
                string timeFormat = "HH:mm";
                string dateFormat = "MMM d";
                string time = DateTime.Now.ToString(timeFormat);
                string date = DateTime.Now.ToString(dateFormat);
                WriteToStatusDisplay($"{Environment.NewLine}{Environment.NewLine}Job complete.", "Job was completed at " + time + " on " + date, Brushes.Black);
            }
        }

        private void numFields_GotFocus(object sender, RoutedEventArgs e)
        {
            rbnFixed.IsChecked = true;
        }

        private void tbFieldLengths_GotFocus(object sender, RoutedEventArgs e)
        {
            rbnFixed.IsChecked = true;
            if (tbFieldLengths.Text == "Enter field lengths separated by commas (e.g. 15,31,24)")
            {
                tbFieldLengths.Text = "";
                tbFieldLengths.Foreground = Brushes.Black;
                tbFieldLengths.FontStyle = FontStyles.Normal;
            }
        }

        // Automatically scrolls the scroll bar when the TextBlock becomes full.
        // This helpful code snippet was obtained from http://stackoverflow.com/a/19315242/816695
        private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            bool autoScroll = true;  // Set or unset autoscroll mode
            if (e.ExtentHeightChange == 0)  // Scroll event detected, but content unchanged
            {
                // Scroll bar is at the bottom, so set autoscroll mode
                if (svStatus.VerticalOffset == svStatus.ScrollableHeight)
                {
                    autoScroll = true;
                }
                else // Scroll bar is not at bottom, so disable autoscroll mode
                {
                    autoScroll = false;
                }
            }

            // Content scroll event : autoscroll eventually
            if ((autoScroll) && (e.ExtentHeightChange != 0))
            {   // Content changed and autoscroll mode set
                // Autoscroll
                svStatus.ScrollToVerticalOffset(svStatus.ExtentHeight);
            }
        }

        private void btnInputFile_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            tblStatusBar.Text = "Load files for conversion to CSV. All selected files must be of the same type and extension.";
        }

        private void MouseLeaveArea(object sender, System.Windows.Input.MouseEventArgs e)
        {
            tblStatusBar.Text = "Relevant tips will appear here as you mouse over the various controls.";
        }

        private void btnSaveFileTo_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            tblStatusBar.Text = "Select folder to save converted CSV files. To keep files organised, consider creating a new output folder.";
        }

        private void btnCreateCsv_MouseEnter(object sender, MouseEventArgs e)
        {
            tblStatusBar.Text = "Begin converting the selected files to CSV. Please make sure the correct file type is chosen.";
        }

        private void tblStatus_MouseEnter(object sender, MouseEventArgs e)
        {
            tblStatusBar.Text = "Status messages from the program will appear here.";
        }

        private void tbFieldLengths_LostFocus(object sender, RoutedEventArgs e)
        {
            // If text box is empty, set its content to the default helpful text.
            if (tbFieldLengths.Text == String.Empty)
            {
                tbFieldLengths.Foreground = Brushes.Gray;
                tbFieldLengths.FontStyle = FontStyles.Italic;
                tbFieldLengths.Text = "Enter field lengths separated by commas (e.g. 15,31,24)";
            }
        }

        private void tbDelimiter_GotFocus(object sender, RoutedEventArgs e)
        {
            rbnDelimited.IsChecked = true;
        }    

        private void lblOutputFilePath_MouseEnter(object sender, MouseEventArgs e)
        {
            if (OutputDirectory == null)
                tblStatusBar.Text = "No output directory specified.";
            else
            {
                lblOutputFilePath.ToolTip = OutputDirectory;
                tblStatusBar.Text = "CSV files will be saved in this location: " + OutputDirectory;
            }
        }

        private void lblSourceFilePath_MouseEnter(object sender, MouseEventArgs e)
        {
            if (inputFiles.Count == 0 || inputFiles == null)
                tblStatusBar.Text = "No input files loaded.";
            else
                tblStatusBar.Text = String.Format("{0} files have been loaded", inputFiles.Count);
        }

        private void cbEncloseWithQualifiers_MouseEnter(object sender, MouseEventArgs e)
        {
            tblStatusBar.Text = "If checked, every field in the CSV file will be enclosed within the specified qualifier.";
        }

        private void cbEncloseWithQualifiers_Checked(object sender, RoutedEventArgs e)
        {
            tbQualifier.IsEnabled = true;
            ValidateInputSuccess();
        }

        private void cbEncloseWithQualifiers_Unchecked(object sender, RoutedEventArgs e)
        {
            tbQualifier.Text = String.Empty;
            ValidateInputSuccess();
        }

        private void tbQualifier_GotFocus(object sender, RoutedEventArgs e)
        {
            cbEncloseWithQualifiers.IsChecked = true;
            ValidateInputSuccess();
        }

        private void FileExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void tbDelimiter_MouseEnter(object sender, MouseEventArgs e)
        {
            tblStatusBar.Text = "Type the delimiting character used to separate individual fields in the file.";
        }

        private void tbFieldLengths_MouseEnter(object sender, MouseEventArgs e)
        {
            tblStatusBar.Text = "Enter the lengths of each field, separating each length with a comma. For example, entering 22,30,44 indicates that the first, second and third fields have lengths 22, 30, and 44 respsectively.";
        }

        private void tbQualifier_MouseEnter(object sender, MouseEventArgs e)
        {
            tblStatusBar.Text = "If the check box is ticked, this character will be used to enclose each field.";
        }

        private void OpenCmdCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void SaveCmdCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void CloseCmdExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            this.Close();
        }

        private void AboutMenuClicked(object sender, RoutedEventArgs e)
        {
            
        }

        private void cbDelimiter_MouseEnter(object sender, MouseEventArgs e)
        {
            tblStatusBar.Text = "Select the delimiting character used to separate individual fields in the file.";
        }

        // Updates the text box tbNumFields accordingly
        private void tbFieldLengths_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (rbnFixed.IsChecked == true)
            {
                // Check if valid sequence of comma-separated whole numbers is entered
                Match match = Regex.Match(tbFieldLengths.Text, @"^\s*[1-9]\d*(?:\s*,\s*[1-9]\d*)*$");
                if (match.Success) // valid field lengths provided 
                {
                    if (inputFiles.Count > 0)
                    {
                        EnableCreateCsvButton();
                    }
                    // Display the field count
                    string[] fieldLengthsArray = tbFieldLengths.Text.Split(','); 
                    lblNumFields.FontWeight = FontWeights.Normal;
                    lblNumFields.Foreground = Brushes.Green;
                    lblNumFields.Content = "Number of fields: " + fieldLengthsArray.Length.ToString();
                    lblNumFields.ToolTip = String.Empty;
                }
                else
                {
                    DisableCreateCsvButton();

                    // Display a red X 
                    lblNumFields.FontWeight = FontWeights.Bold;
                    lblNumFields.Foreground = Brushes.Red;
                    lblNumFields.FontSize = 16;
                    lblNumFields.Content = "X";
                    lblNumFields.ToolTip = "Field lengths must be given as comma-separated whole numbers.";
                }
            }
        }

        /// <summary>
        /// Disables the "Create CSV" button and changes its appearance
        /// </summary>
        /// <param name="tooltip">The tooltip text for the button.</param>
        private void DisableCreateCsvButton()
        {
            btnCreateCsv.IsEnabled = false;
            btnCreateCsv.Foreground = Brushes.Red;
            btnCreateCsv.Opacity = 0.5;
        }

        /// <summary>
        /// Enables the "Create CSV" button and changes its appearance
        /// </summary>
        /// <param name="tooltip">The tooltip text for the button.</param>
        private void EnableCreateCsvButton()
        {
            btnCreateCsv.IsEnabled = true;
            btnCreateCsv.Foreground = Brushes.MediumSeaGreen;
            btnCreateCsv.Opacity = 1.0;
        }

        private void cbDelimiterItemSelected(object sender, RoutedEventArgs e)
        {
            if (ValidateInputSuccess())
                EnableCreateCsvButton();
            else
                DisableCreateCsvButton();
        }

        private void tbDelimiter_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (ValidateInputSuccess())
                EnableCreateCsvButton();
            else
                DisableCreateCsvButton();
        }

        private void cbEncloseWithQualifiers_Click(object sender, RoutedEventArgs e)
        {
            if (ValidateInputSuccess())
                EnableCreateCsvButton();
            else
                DisableCreateCsvButton();
        }
        
        private void tbQualifier_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrWhiteSpace(tbQualifier.Text))
                cbEncloseWithQualifiers.IsChecked = false;
            if (ValidateInputSuccess())
                EnableCreateCsvButton();
            else
                DisableCreateCsvButton();
        }

        /// <summary>
        /// Opens a file browser dialog, reads fields from the selected layout file, and updates the text box tbFieldLengths accordingly.
        /// </summary>
        private void btnLoadLayoutFile_Click(object sender, RoutedEventArgs e)
        {
            // Show an open file dialog
            var dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".lay";
            dlg.Filter = "Layout Files|*.lay";
            var result = dlg.ShowDialog();
            
            // If user opens a file, read field lengths and put them into the text box tbFieldLengths.
            if (result == true)
            {
                rbnFixed.IsChecked = true;
                tbFieldLengths.FontStyle = FontStyles.Normal;
                tbFieldLengths.Foreground = Brushes.Black;
                List<string> fieldLengths = GetFieldLengthsFromLayoutFile(dlg.FileName);
                string commaSeparatedFieldLengths = String.Join(",", fieldLengths.ToArray());
                tbFieldLengths.Text = commaSeparatedFieldLengths;
            }
        }

        private void rbnFixed_Unchecked(object sender, RoutedEventArgs e)
        {
            lblNumFields.Content = String.Empty;
        }

        #endregion
    }
}
