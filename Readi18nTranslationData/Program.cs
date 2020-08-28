using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace Readi18nTranslationData

{
    class Program
    {
        static void Main(string[] args)
        {
            //String fileName = @"C:\Users\JCK0412\source\repos\ExcelParser\Readi18nTranslationData\data\HQ-Internationalization-v2-Finnish.xlsx";
            String fileName = @"C:\Users\JCK0412\source\repos\ExcelParser\Readi18nTranslationData\data\sample.xlsx";
            //string sheetName = "Locations";

            //get sheet names dynamically
            //    https://docs.microsoft.com/en-us/office/open-xml/how-to-retrieve-a-list-of-the-worksheets-in-a-spreadsheet





     
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                { 
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    //workbookPart.WorksheetParts.Count()

                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    Worksheet sheet = worksheetPart.Worksheet;

                    var cells = sheet.Descendants<Cell>();
                    var rows = sheet.Descendants<Row>();

                    //  Console.WriteLine("Row count = {0}", rows.LongCount());
                    //  Console.WriteLine("Cell count = {0}", cells.LongCount());

                    bool loadData = false;
                    char englishTextColumn = 'C';
                    char tagColumn = 'E';
                    List<TranslationKey> translationKeys = new List<TranslationKey>();
                    int translationKeyID = 0;

                    foreach (Row row in rows)
                    {
                        string englishText = "";
                        string tag = "";
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(cell.CellValue.Text);
                                string cellValue = sst.ChildElements[ssid].InnerText;


                                if (loadData)
                                {
                                    string cellRef = cell.CellReference.ToString();
                                    char columnLetter = cellRef[0];
                                    
                                    if (columnLetter == englishTextColumn)
                                    {
                                        englishText = cellValue;   
                                    }

                                    if (columnLetter == tagColumn)
                                    {
                                        tag = cellValue;
                                    }
                                }

                                // do I need to check for "Screen" or can I just run the loader as-is;
                                if (cellValue == "Screen")
                                {
                                    loadData = true;
                                    break;
                                }
                            }
                        }

                        if (String.IsNullOrEmpty(tag) == false)
                        {
                            translationKeyID++;
                            TranslationKey translationKey = new TranslationKey
                            {
                                TranslationKeyID = translationKeyID,
                                Tag = tag
                            };
                            translationKeys.Add(translationKey);
                        };
                    }
                    Program.GenerateSQLTranslationKey(translationKeys);
                }
                Console.ReadLine();
            }

        }

        public static void GenerateSQLTranslationKey(List<TranslationKey> translationKeys)
        {
            Console.WriteLine ("gimme me some sql");

            //todo: create object instead of list of strings;

            foreach (var translationKey in translationKeys)
            {
                string sqlInsert = String.Format("INSERT #TranslationKey (TranslationKeyID, Tag) VALUES ({0}, '{1}')",translationKey.TranslationKeyID.ToString(), translationKey.Tag);
                Console.WriteLine(sqlInsert);
            }
        }
    }
}

//INSERT #TranslationKey (TranslationKeyID, Tag) VALUES (1, 'LoginModule_Title')
