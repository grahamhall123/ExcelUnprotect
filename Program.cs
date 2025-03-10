using System;
using CommandLine;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace ExcelUnprotect
{
    class Program
    {
        static void Main(string[] args)
        {
            // Remove any temporary folders we may have previously created.
            tidyUp();

            var options = new Options();

            Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(o =>
                {
                    // Check the path.
                    if (File.Exists(o.FilePath))
                    {
                        string ext = Path.GetExtension(o.FilePath);

                        // Check Extension.
                        if (ext == ".xls" || ext == ".xlsx" || ext == ".xlsm")
                        {
                            // Is the file in use?
                            if (!IsFileLocked(o.FilePath))
                            {
                                // Set output file.
                                string outputFile = o.OutputFilePath;
                                if (o.OutputFilePath == null) outputFile = o.FilePath.Replace(ext, "") + "-unprotected" + ext;

                                Console.WriteLine("\r\nChecking spreadsheet...\r\n");

                                if (unZip(o.FilePath))
                                {
                                    bool protectedWorkbook = hasProtectedWorkbook();
                                    bool protectedSheet = hasProtectedSheets();

                                    // If we have protection.
                                    if (protectedWorkbook || protectedSheet)
                                    {
                                        Console.WriteLine("\r\nPress y to remove protection or any other key to cancel");

                                        ConsoleKeyInfo keypress = Console.ReadKey();

                                        Console.WriteLine("\r\n");

                                        if (keypress.Key.ToString() == "Y")
                                        {
                                            bool removed = false;

                                            if (protectedWorkbook)
                                            {
                                                removeWorkbookProtection();

                                                removed = true;
                                            }

                                            if (protectedSheet)
                                            {
                                                removeSheetProtection();

                                                removed = true;
                                            }

                                            // If we've removed.
                                            if (removed)
                                            {
                                                repackage(outputFile);
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine("\r\nNo action taken");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("\r\nSpreadsheet doesn't look as though it has protection");
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("\r\nFile is already open. Make sure you close it in Excel first!");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Only .xls, .xlsx and xlsm files supported");
                        }
                    }
                    else
                    {
                        Console.WriteLine("\r\nCouldn't find file " + o.FilePath);
                    }
                });
            
            // Remove any temporary folders.
            tidyUp();
        }


        private static bool unZip(string zipPath)
        {
            string extractPath = "unprot-tmp";

            // Create a new hidden directory.
            if (!Directory.Exists(extractPath))
            {
                DirectoryInfo di = Directory.CreateDirectory(extractPath);
                di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
            }

            // Extract.
            try
            {
                ZipFile.ExtractToDirectory(zipPath, extractPath);
                return true;
            }
            catch
            {
                Console.WriteLine("Couldn't open. Make sure it's a newer type of Excel document");
                Console.WriteLine("Tool can't remove spreadsheet encryption");
                return false;
            }
        }


        // Check the workbook file for protection.
        private static bool hasProtectedWorkbook()
        {
            // Check the workbook.xml file.
            string workbookPath = "unprot-tmp/xl/workbook.xml";

            if (File.Exists(workbookPath))
            {
                string text = File.ReadAllText(workbookPath);

                if (text.Contains("<workbookProtection"))
                {
                    Console.WriteLine("# Workbook is protected");
                    return true;
                }
            }

            return false;
        }


        // Check all sheets for protection.
        private static bool hasProtectedSheets()
        {
            string worksheetsPath = "unprot-tmp/xl/worksheets/";
            bool isProtected = false;

            if (Directory.Exists(worksheetsPath))
            {
                int fileCount = Directory.GetFiles(worksheetsPath, "sheet*.xml", SearchOption.TopDirectoryOnly).Length;

                if (fileCount > 0)
                {
                    // Check through each sheet.
                    for (int i = 1; i < fileCount+1; i++)
                    {
                        string filename = worksheetsPath + "sheet" + i + ".xml";

                        if (File.Exists(filename))
                        {
                            // Open the file as text.
                            string text = File.ReadAllText(filename);

                            // Check for tags.
                            if (text.Contains("<protectedRanges") || text.Contains("<protectedRange") || text.Contains("<sheetProtection"))
                            {
                                Console.WriteLine("# Sheet " + i + " is protected");
                                isProtected = true;
                            }
                        }
                    }
                }
            }

            return isProtected;
        }


        private static void removeWorkbookProtection()
        {
            // Check the workbook.xml file.
            string workbookPath = "unprot-tmp/xl/workbook.xml";

            if (File.Exists(workbookPath))
            {
                var document = new XmlDocument();

                document.Load(workbookPath);
                   
                // Because we have namespaces.
                var nsmgr = new XmlNamespaceManager(document.NameTable);
                nsmgr.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                // Get a list of nodes with the tag to be removed.
                XmlNodeList nodes = document.DocumentElement.SelectNodes("//x:workbookProtection", nsmgr);

                if (nodes == null) return;

                foreach (XmlNode node in nodes)
                {
                    document.DocumentElement.RemoveChild(node);
                }

                Console.WriteLine("# Removed workbook protection");

                document.Save(workbookPath);
            }
        }


        private static void removeSheetProtection()
        {
            string worksheetsPath = "unprot-tmp/xl/worksheets/";

            if (Directory.Exists(worksheetsPath))
            {
                int fileCount = Directory.GetFiles(worksheetsPath, "sheet*.xml", SearchOption.TopDirectoryOnly).Length;

                if (fileCount > 0)
                {
                    // Check through each sheet.
                    for (int i = 1; i < fileCount + 1; i++)
                    {
                        string filename = worksheetsPath + "sheet" + i + ".xml";

                        if (File.Exists(filename))
                        {
                            // Open the file as text.
                            string text = File.ReadAllText(filename);

                            // Check for tags.
                            if (text.Contains("<protectedRanges") || text.Contains("<sheetProtection"))
                            {
                                var document = new XmlDocument();

                                document.Load(filename);

                                // Because we have namespaces.
                                var nsmgr = new XmlNamespaceManager(document.NameTable);
                                nsmgr.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                                // Get a list of nodes with the tag to be removed.
                                XmlNodeList nodes = document.DocumentElement.SelectNodes("//x:protectedRanges", nsmgr);

                                if (nodes != null)
                                {
                                     foreach (XmlNode node in nodes)
                                    {
                                        document.DocumentElement.RemoveChild(node);
                                    }
                                }

                                // Get a list of nodes with the tag to be removed.
                                nodes = document.DocumentElement.SelectNodes("//x:sheetProtection", nsmgr);

                                if (nodes != null)
                                {
                                    foreach (XmlNode node in nodes)
                                    {
                                        document.DocumentElement.RemoveChild(node);
                                    }
                                }

                                Console.WriteLine("# Removed protection on sheet " + i);

                                document.Save(filename);
                            }
                        }
                    }
                }
            }
        }


        private static void repackage(string filePath)
        {
            string startPath = "unprot-tmp";

            // Check if the file exists first.
            if (File.Exists(filePath))
            {
                // Remove.
                File.Delete(filePath);
            }

            ZipFile.CreateFromDirectory(startPath, filePath);

            Console.WriteLine("\r\n" + filePath + " created");
        }


        // Remove any temporary directories.
        private static void tidyUp()
        {
            string extractPath = "unprot-tmp";

            if (Directory.Exists(extractPath))
            {
                Directory.Delete(extractPath, true);
            }
        }


        protected static bool IsFileLocked(string filePath)
        {
            try
            {
                using (Stream stream = new FileStream(filePath, FileMode.Open))
                {
                    return false;
                }
            }
            catch
            {
                return true;
            }
        }

        // Check Exists.
        // Check XLSX.
        // Try to Unpack to temporary directory.
        // Check Spreadsheet Protection.
        // Check Sheet Protection.
        // Show any protected Spreadsheet / sheets.
        // Ask for yes / no.
        // Unprotect Spreadsheet.
        // Unprotect Sheets.
        // Repack to Output filename / filename-unprotected.xlsx.
        // Give Results.
    }

    class Options
    {
        [Option('f', "file", Required = true, HelpText = "File path of Excel document.")]
        public String FilePath { get; set; }

        [Option('o', "outputfile", Required = false, HelpText = "File path of the output document.")]
        public String OutputFilePath { get; set; }
    }
}
