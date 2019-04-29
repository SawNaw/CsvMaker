using JonathanWood.ReadWriteCsv;
using System.Data.OleDb;
using System.IO;
using System;

namespace CsvMaker
{
    /// <summary>
    /// Contains methods that processes various types of data files and converts them to CSV.
    /// </summary>
    public static class CsvConverter
    {
        /// <summary>
        /// Represents the result of a CSV conversion.
        /// </summary>
        public class CsvConverterResult
        {  
            public string Message { get; set; }
            public int LineNumber { get; set; }
            public string FileName { get; set; }

            /// <summary>
            /// Holds information about the result of a CSV conversion operation. Use this overload to return an error.
            /// </summary>
            /// <param name="message">The message regarding the result of the operation.</param>
            /// <param name="lineNumber">The line number in the file at which the result occurred.</param>
            /// <param name="fileName">The full path of the file in which the error occurred.</param>
            public CsvConverterResult(string message, int lineNumber, string fileName)
            {
                Message = message;
                LineNumber = lineNumber;
                FileName = fileName;
            }
        }
        
        /// <summary>
        /// Reads a delimited data file and produces a CSV file from it.
        /// </summary>
        /// <param name="delimiter">The delimiting character used in the delimited data file.</param>
        /// <param name="inputFilePath">The full path of the file to be converted.</param>
        /// <param name="outputFilePath">The full path of the CSV file to be output.</param>
        /// <param name="qualifier">The qualifier to surround each field with.</param>
        /// <returns>
        /// If errors were encountered, returns a CsvConverterResult which contains information about those errors.
        /// If no errors were encountered, returns null.
        /// </returns>
        public static CsvConverterResult WriteCsvRowsToFile(char delimiter, string inputFilePath, string outputFilePath, string qualifier = "")
        {
            CsvConverterResult error = null; // any errors encountered will be added to this list

            using (CsvFileReader reader = new CsvFileReader(inputFilePath, delimiter)) // create a CsvFileReader that knows the input path and delimiter
            using (CsvFileWriter writer = new CsvFileWriter(outputFilePath)) // create a CsvFileWriter that knows which path to write to
            {
                CsvRow firstRow = new CsvRow();  // the first row 
                CsvRow firstOutputRow = new CsvRow();  // the first row, whitespaced trimmed for output
                int currentLineNumber = 0; // the line number of the line that was most recently read

                // Begin by reading the first row and writing it to the CSV file.
                reader.ReadRow(firstRow); 
                currentLineNumber++;
                
                foreach (string s in firstRow)
                {
                    firstOutputRow.Add(s.Trim());
                }

                // Write the CsvRow to the file.
                writer.WriteRow(firstOutputRow, qualifier);

                // Now read subsequent rows and compare them to the first row. 
                // If both rows have an equal number of fields, add to output file. If not, signal error.
                CsvRow inputRow = new CsvRow();
                while (reader.ReadRow(inputRow))
                {
                    currentLineNumber++;
                    if (inputRow.Count == firstRow.Count) // Check if the newly read row has the same number of fields as the first row
                    {
                        CsvRow outputRow = new CsvRow();
                        foreach (string s in inputRow)
                        {
                            outputRow.Add(s.Trim());
                        }

                        // Write the CsvRow to the file
                        writer.WriteRow(outputRow, qualifier);
                    }
                    else  // number of fields in current row is different from the first row, possible error in CSV file
                    {
                        string msg = String.Format("This line has a different number of fields than previous lines. Conversion cancelled.", 
                            currentLineNumber);
                        error = new CsvConverterResult(msg, currentLineNumber, inputFilePath);
                        break;
                    }
                }
            }
            return error;
        }

        /// <summary>
        /// Reads a fixed width data file and produces a CSV file from it. Discards any newline characters occurring inside fields.
        /// </summary>
        /// <param name="fieldWidths">An array containing the lengths of each field.</param>
        /// <param name="inputFilePath">The full path of the file to be converted</param>
        /// <param name="outputFilePath">The output file path of the CSV file to be created.</param>
        /// <param name="qualifier">The qualifier to surround each field with.</param>
        /// <returns> Returns 0 if no exceptions occurred. Returns the problematic line number of the text file if a parsing error occurs.</returns>
        public static long WriteCsvRowsToFile(int[] fieldWidths, string inputFilePath, string outputFilePath, string qualifier = "")
        {
            long errorLine = 0;  // the line number at which an error occurred

            using (StreamReader reader = new StreamReader(inputFilePath))
            using (CsvFileWriter writer = new CsvFileWriter(outputFilePath))
            {
                do
                {
                    CsvRow row = new CsvRow();  // CSV row to construct

                    // For every field length
                    for (int i = 0; i < fieldWidths.Length; i++)
                    {
                        // Read the field into a char array.
                        char[] field = new char[fieldWidths[i]];
                        reader.ReadBlock(field, 0, fieldWidths[i]);

                        // If the field has two characters and they are "\r\n", replace with blank characters.
                        if (field.Length == 2 && field[0] == '\r' && field[1] == '\n')
                        {
                            field[0] = ' ';
                            field[1] = ' ';
                        }
                        row.Add(new string(field).Trim());  // Add this field to the CsvRow
                    }
                    writer.WriteRow(row, qualifier); // Write the CsvRow to the file

                } while (!reader.EndOfStream);
            }
            return errorLine;
        }

        /// <summary>
        /// Reads an Excel file and produces a CSV file from it.
        /// </summary>
        /// <param name="command">The OLEDB command to connect to an Excel file.</param>
        /// <param name="outputFilePath">The desired output file path of the CSV file.</param>
        /// <param name="qualifier">The qualifier to surround each field with.</param>
        /// <remarks>
        /// This overload does not require an input file path passed to it, since that information is 
        /// present in the OleDBCommand used to connect to the Excel file.
        /// </remarks>
        public static void WriteCsvRowsToFile(OleDbCommand command, string outputFilePath, string qualifier = "")
        {
            using (OleDbDataReader dr = command.ExecuteReader())
            using (CsvFileWriter writer = new CsvFileWriter(outputFilePath))
            {
                while (dr.Read()) // read the next record
                {
                    CsvRow row = new CsvRow();
                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        if (!dr.IsDBNull(i)) // field is not empty, so add its text to the CsvRow
                        {
                            //string field = dr.GetString(i).Trim();
                            string field = dr[i].ToString().Trim();
                            row.Add(field);
                        }
                        else // field is empty, so add a blank string to the CsvRow
                        {
                            row.Add("");
                        }
                    }
                    // Write the CsvRow to the file
                    writer.WriteRow(row, qualifier);
                }
            }
        }
    }
}
