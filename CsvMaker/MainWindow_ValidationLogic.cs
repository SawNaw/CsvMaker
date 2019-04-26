using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace CsvMaker
{
    /// <summary>
    /// Validation logic for MainWindow.xaml.cs
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Checks whether all files in the string array have the same extension.
        /// </summary>
        /// <param name="files">The string array of filenames</param>
        /// <returns>True if all filenames have the same extension. False if otherwise.</returns>
        private bool AllFilesHaveSameExtension(string[] files)
        {
            bool haveSameExtensions = true;
            string extensionOfFirstFile = Path.GetExtension(files.ElementAt<string>(0));
            foreach (string fileName in files)
            {
                if (Path.GetExtension(fileName) != extensionOfFirstFile)
                {
                    haveSameExtensions = false;
                    break;
                }
            }
            return haveSameExtensions;
        }

        /// <summary>
        /// Returns true if the field lengths in tbFieldLengths are entered correctly
        /// </summary>
        /// <returns></returns>
        private bool FieldLengthsEnteredCorrectly()
        {
            bool isCorrect = false;
            string pattern = @"^\s*[1-9]\d*(?:\s*,\s*[1-9]\d*)*$";  // positive whole numbers separated by commas, allow spaces
            Match match = Regex.Match(tbFieldLengths.Text, pattern);
            if (match.Success)
                isCorrect = true;

            return isCorrect;
        }

        /// <summary>
        /// Determines if the controls on the window are in a state that is ready for processing.
        /// Disables the "Create CSV" button if validation fails. Enables it if validation succeeds.
        /// </summary>
        /// <returns>True if input validation was a success, and false otherwise.</returns>
        private bool ValidateInputSuccess()
        {
            bool fileTypeGroupBoxValidationSuccess = false; 
            bool optionsGroupBoxValidationSuccess = false;

            // If no input files were selected, immediately fail validation.
            if (inputFiles.Count() == 0)  
                return false;

            // If file is specified as delimited type, validate accordingly.
            if (rbnDelimited.IsChecked == true)  // delimited file specified
            {
                if (cbiPipe.IsSelected || cbiSemicolon.IsSelected || cbiTab.IsSelected)  // delimiter is selected
                    fileTypeGroupBoxValidationSuccess = true;
                if (cbiOther.IsSelected && (tbDelimiter.Text != String.Empty))  // "Other" delimiter is selected and delimiting character is specified
                {
                    char delimiter = tbDelimiter.Text[0];
                    if (Char.IsLetterOrDigit(delimiter))
                    {
                        string caption = "Unexpected delimiter type!";
                        string msg = "You specified a letter or a digit as the delimiter. Are you sure you want to do this?";
                        MessageBoxResult result = MessageBox.Show(msg, caption, MessageBoxButton.YesNo, MessageBoxImage.Warning);
                        if (result == MessageBoxResult.Yes)
                            fileTypeGroupBoxValidationSuccess = true;
                        else
                            tbDelimiter.Text = String.Empty;
                    }
                    else  // delimiter is not a letter or number, so accept it
                        fileTypeGroupBoxValidationSuccess = true;
                }
                else  // delimiter not specified
                {
                    
                }
            }

            // If file is specified as having fixed length data, validate accordingly.
            else if (rbnFixed.IsChecked == true) // Fixed length file specified.
            {
                // Check if valid sequence of comma-separated whole numbers is entered
                Match match = Regex.Match(tbFieldLengths.Text, @"^\s*[1-9]\d*(?:\s*,\s*[1-9]\d*)*$");
                if (match.Success)
                    fileTypeGroupBoxValidationSuccess = true;
            }

            // If file is specified as being an Excel file, validate accordingly.
            else if (rbnExcelXls.IsChecked == true || rbnExcelXlsx.IsChecked == true) // .XLS or .XLSX file expected
            {
                // Check if selected files have the .xls or .xlsx extension
                foreach (string file in inputFiles)
                {
                    string extension = Path.GetExtension(file).ToLower();
                    if (extension != ".xls" && extension != ".xlsx")
                    {
                        string msg = $"{Path.GetFileName(file)} does not seem to be an Excel file.{Environment.NewLine}";
                        string tooltipText = "You specified that you will be loading Excel files, but one or more files have an extension other than .xls or .xlsx";
                        WriteToStatusDisplay(msg, tooltipText, Brushes.Crimson, Brushes.White);
                    }
                    else
                        fileTypeGroupBoxValidationSuccess = true;
                } 
            }
            else  // if this block is reached, then an unknown radio button was selected. The user shouldn't be seeing this!
            {
                string msg = "Fatal error: Unexpected RadioButton parameter received in ValidateInput(). Please inform the developer.";
                string caption = "Unexpected RadioButton selected!";
                MessageBox.Show(msg, caption, MessageBoxButton.OK, MessageBoxImage.Stop);
            }

            // If the "enclose with qualifier" check box is checked, but a qualifier isn't specified, fail validation.
            if (cbEncloseWithQualifiers.IsChecked == true && String.IsNullOrWhiteSpace(tbQualifier.Text))
                DisableCreateCsvButton();
            else
                optionsGroupBoxValidationSuccess = true;

            return (fileTypeGroupBoxValidationSuccess && optionsGroupBoxValidationSuccess);
        }
    }
}
