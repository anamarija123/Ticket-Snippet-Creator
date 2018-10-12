using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks; 

namespace TicketJsonConverter
{

    /*
     TicketJsonConverter is tool for making snippets in ticket language.
     TJ Converter generates any file format you want in text structure for load snippets in VisualStudioCode.
         */
    class Program
    {
        private const int PARAMETER_MISSING = 1;
        private const int INVALID_CAST = 2;
        private const int INVALID_DATA = 3;
        private const int EXCEL_SOURCE_FILE_ERROR = 4;
        private const int ERROR = 5;

        static int Main(string[] args)
        {
          
            //checking for input parameters
            if (args.Length < 1)
            {
                PrintUsageDetails("Parametars are missing!");
                return PARAMETER_MISSING;
            }
            string fileName;
            string destinationFileName;
            List<Tuple<string, string, string>> result = new List<Tuple<string, string, string>>();

            try
            {
                fileName = args[0];
                destinationFileName = args[1];
            }
            catch (System.IndexOutOfRangeException ex)
            {
                Console.WriteLine("Invalid path!");
                Console.WriteLine($"Description: {ex.Message}");
                return PARAMETER_MISSING;
            }

            if (!File.Exists(fileName))
            {
                Console.WriteLine("File not found.");
                return 2;
            }

            try
            {
                //read from excel file and put required elements in list
                
                using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
                {

                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read() && reader.Name.Equals("Bull"))
                            {
                                try
                                {
                                    var tag = reader.GetString(0);
                                    var name = reader.GetString(1);
                                    var description = reader.GetString(2);

                                    // Skip header
                                    if (tag.ToLower().Equals("tag"))
                                    {
                                        continue;
                                    }
                                    // set description and name if is not given in file
                                    if (description == null || name == null || description == "?" || description == "_")
                                    {
                                        description = tag;
                                        name = tag;
                                    }
                                    // if there is any new lines in description, take only first line before new line
                                    else
                                    {
                                        string[] parts = description.Split('\n');
                                        description = parts[0];
                                    }

                                    result.Add(new Tuple<string, string, string>(tag, name, description));
                                }
                                catch (InvalidDataException ex)
                                {
                                    Console.WriteLine("Invalid data, check your input file");
                                    Console.WriteLine($"Description: {ex.Message}");
                                    return INVALID_DATA;
                                }
                                catch (InvalidCastException ex)
                                {
                                    Console.WriteLine("Invalid cast, check your input file");
                                    Console.WriteLine($"Description: {ex.Message}");
                                    return INVALID_CAST;
                                }
                                catch (ArgumentException ex)
                                {
                                    Console.WriteLine($"Error: invalid input parameter");
                                    Console.WriteLine($"Description: {ex.Message}");
                                    return PARAMETER_MISSING;
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error: input parameters error");
                                    Console.WriteLine($"Description: {ex.Message}");
                                    return PARAMETER_MISSING;
                                }
                            }

                        } while (reader.NextResult());

                    }
                }
            }
            catch (ExcelDataReader.Exceptions.HeaderException ex)
            {
                Console.WriteLine($"Error: invalid source file");
                Console.WriteLine($"Description: {ex.Message}");
                return EXCEL_SOURCE_FILE_ERROR;
            }
            catch (ExcelDataReader.Exceptions.ExcelReaderException ex)
            {
                Console.WriteLine($"Error: reading file error");
                Console.WriteLine($"Description: {ex.Message}");
                return EXCEL_SOURCE_FILE_ERROR;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: error");
                Console.WriteLine($"Description: {ex.Message}");
                return ERROR;
            }
            
            //write how many items are readed from file
            Console.WriteLine(result.Count);

            //making structure to write in file
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("{");
            result.ForEach(i => sb.AppendLine(formatItem(i)));
            sb.AppendLine("}");

            //write in file
            File.WriteAllText(destinationFileName, sb.ToString());

            return 0;
        }
        //making format to write in file
        private static string formatItem(Tuple<string, string, string> item)
        {
            return String.Format(
                "\t\"{0}\": {{\n\t\t\"prefix\": \"#{1}\",\n\t\t\"body\": \"#{1}\",\n\t\t\"description\": \"{2}\"\n\t}},",
                item.Item2, item.Item1, item.Item3);
        }

        static void PrintUsageDetails(string errorDescription)
        {
            Console.WriteLine("Error: " + errorDescription);
            Console.WriteLine();
            Console.WriteLine("TicketJsonConverter.exe: Use to generate snippets for ticket language");
            Console.WriteLine();
            Console.WriteLine("Usage:\n   TicketJsonConverter.exe /source: sourceFile path  /destination: destination path with file name");
            Console.WriteLine("   /sourceFile path example: C:\\Users\anama\\Desktop\\sadrzaj\\tags.xls");
            Console.WriteLine("   sourceFile must have one of this extensions: .xls, .xlsx, .xlsm, .xlsb, .xltx, .xltm, .xlt, .xml");
            Console.WriteLine();
            Console.WriteLine("   /destination file path example: C:\\Users\anama\\Desktop\\NewFolder1\\ticket.json");
            Console.WriteLine("   path to the new file where we want to write snippets, it can have any extension you want");
            Console.WriteLine();
        }
    }
}
