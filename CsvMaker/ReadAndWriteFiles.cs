// Contains private methods that are executed immediately after 
// the type of input file has been specified by the user

using JonathanWood.ReadWriteCsv;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace NewApp
{
    public partial class MultiFormatTextFileReader
    {
        private static void ReadExcelFileAndOutputToCsv()
        {
            string inputFilePath = null;
            string outputFilePath = null;
            string connString = null;
            
            // Get the file path. This method also returns the type of Excel file (.xls or .xlsx)
            string excelFileType = PromptForExcelFileParameters(out inputFilePath, out outputFilePath);

            // Form connection string for specified Excel 2003 file
            switch (excelFileType)
            {
                case ".xls":
                    connString = String.Format(@"Provider=Microsoft.Jet.OLEDB.4.0; 
                                        Data Source={0}; 
                                        Extended Properties='Excel 8.0; HDR=Yes;'"
                                        , inputFilePath);
                    break;
                case ".xlsx":
                    connString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;
                                        Data Source={0};
                                        Extended Properties='Excel 12.0 Xml; HDR=YES; IMEX=1;'" // IMEX=1 treats all data as text
                                        , inputFilePath);
                    break;
            }

            // Create new OleDB connection 
            using (OleDbConnection connection = new OleDbConnection(connString))
            {
                connection.Open(); 
                string queryString = "SELECT * FROM [Sheet1$]";
                OleDbCommand command = new OleDbCommand(queryString, connection);

                // Read fields from the Excel file and add them to CsvRow row
                using (OleDbDataReader dr = command.ExecuteReader())
                using (CsvFileWriter writer = new CsvFileWriter(outputFilePath))
                {
                    while (dr.Read()) // read the next record
                    {
                        CsvRow row = new CsvRow();
                        for (int i = 0; i < dr.FieldCount; i++)
                        {
                            if (!dr.IsDBNull(i)) // if field is not empty, so add its text to the CsvRow
                            {
                                string field = dr.GetString(i).Trim();
                                row.Add(field);
                            }
                            else // if field is empty, so add a blank string to the CsvRow
                            {
                                row.Add("");
                            }
                        }
                        writer.WriteRow(row); // write the CsvRow to the CSV file, inserting commas as delimiters
                    }
                }
                Console.WriteLine("CSV file created at {0}", outputFilePath);
            }
        }

        private static void ReadDelimitedFileAndOutputToCsv()
        {
            string inputPath = null;
            string outputPath = null;
            char delimiter = ',';
            PromptForDelimitedFileParameters(out inputPath, out delimiter, out outputPath);

            using (CsvFileReader reader = new CsvFileReader(inputPath, delimiter))
            using (CsvFileWriter writer = new CsvFileWriter(outputPath))
            {
                CsvRow inputRow = new CsvRow();
                while (reader.ReadRow(inputRow))
                {
                    CsvRow outputRow = new CsvRow();
                    foreach (string s in inputRow)
                    {
                        outputRow.Add(s.Trim());
                    }
                    writer.WriteRow(outputRow); 
                }
            }
            Console.WriteLine("CSV file created at {0}", outputPath);
        }

        // Read the fixed length record file and write to CSV
        private static void ReadFixedLengthFileAndOutputToCsv()
        {
            string inputFilePath = null; 
            string outputFilePath = null; 
            int numFields = 0;
            List<int> fieldWidths = new List<int>(); // length of each field
            PromptForFixedWidthParameters(out inputFilePath, out numFields, fieldWidths, out outputFilePath);

            // Read rows from the fixed length record file
            TextFieldParser parser = new TextFieldParser(inputFilePath);
            parser.TextFieldType = FieldType.FixedWidth;
            parser.SetFieldWidths(fieldWidths.ToArray());

            Console.WriteLine(String.Format("Writing to CSV file {0}", outputFilePath));
            using (CsvFileWriter writer = new CsvFileWriter(outputFilePath))
            {
                while (!parser.EndOfData)
                {
                    // Read all fields on the current line, having knowledge of their lengths
                    string[] fields = parser.ReadFields();

                    // Add each field to a CsvRow object
                    CsvRow row = new CsvRow();
                    foreach (string s in fields)
                    {
                        row.Add(s);
                    }

                    // Write that CsvRow to the output file
                    writer.WriteRow(row);
                }
            }
        }
    }
}
