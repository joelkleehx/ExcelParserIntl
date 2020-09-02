using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Office.Word;

namespace Readi18nTranslationData
{
    class Program
    {

        public static int translationKeyID { get; set; }

        static void Main(string[] args)
        {
            List<string> insertSQL = new List<string>();

            String fileName = @"C:\Users\JCK0412\source\repos\ExcelParser\Readi18nTranslationData\data\HQ-Internationalization-v2-Finnish.xlsx";
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    List<WorksheetPart> worksheetParts = workbookPart.WorksheetParts.ToList();

                    SQLStatement sqlStatement = new SQLStatement();
                    sqlStatement.StatementLines = null;
                
                    string[] initialSQL = File.ReadAllLines(String.Format(@"{0}\sqlTemplates\TranslationKeyInit.txt", Environment.CurrentDirectory));
                    string[] loadSQL = File.ReadAllLines(String.Format(@"{0}\sqlTemplates\TranslationKeyLoadRecords.txt", Environment.CurrentDirectory));

                    foreach (WorksheetPart worksheetPart in worksheetParts)
                    {
                        ProcessWorksheet(worksheetPart, sstpart, sst, insertSQL);
                    }

                    string path = Path.Combine(".", "OutputData");
                    string fullTranslationKeyFile = Path.Combine(path, "661-TranslationKey.sql");

                    Directory.CreateDirectory(path);

                    File.WriteAllLines(fullTranslationKeyFile, initialSQL);
                    File.AppendAllLines(fullTranslationKeyFile, insertSQL);
                    File.AppendAllLines(fullTranslationKeyFile, loadSQL);

                
                }
            }
        }

        public static void ProcessWorksheet(WorksheetPart worksheetPart, SharedStringTablePart sstpart, SharedStringTable sst, List<string> insertSQL)
        {
            Worksheet sheet = worksheetPart.Worksheet;
            var rows = sheet.Descendants<Row>();
           
            bool loadData = false;
            char englishTextColumn = 'C';
            char tagColumn = 'E';
            List<Translation> translations = new List<Translation>();
            List<TranslationKey> translationKeys = new List<TranslationKey>();
           
            int translationID = 0;

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
                    if (translationKeys.Any(x => x.Tag == tag) == false)
                    {
                        translationKeyID++;
                        TranslationKey translationKey = new TranslationKey
                        {
                            TranslationKeyID = translationKeyID,
                            Tag = tag
                        };
                        translationKeys.Add(translationKey);
                    }

                    translationID++;
                    Translation translation = new Translation
                    {
                        TranslationLanguageID = (int)Language.USEnglish,
                        TranslationKeyID = translationKeyID,
                        TranslatedText = englishText,
                    };
                    translations.Add(translation);

                };
            }

            List<string> translationKeySQL = Program.GenerateSQLTranslationKey(translationKeys);
            foreach (string insertStatement in translationKeySQL)
            {
                insertSQL.Add(insertStatement);
            }




        }

        public static void LoadTranslationKey(List<TranslationKey> translationKeys)
        {
          
            List<string> insertSQL = Program.GenerateSQLTranslationKey(translationKeys);



        }

        public static void LoadTranslation(List<Translation> translations)
        {
            string[] initialSQL = File.ReadAllLines(String.Format(@"{0}\sqlTemplates\TranslationInit.txt", Environment.CurrentDirectory));
            List<string> insertSQL = Program.GenerateSQLTranslation(translations);
            string[] loadSQL = File.ReadAllLines(String.Format(@"{0}\sqlTemplates\TranslationLoadRecords.txt", Environment.CurrentDirectory));

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\temp\662-Translation.sql"))
            {
                foreach (string line in initialSQL)
                {
                    file.WriteLine(line);
                }

                foreach (string line in insertSQL)
                {
                    file.WriteLine(line);
                }

                foreach (string line in loadSQL)
                {
                    file.WriteLine(line);
                }
            }
        }

        public static List<string> GenerateSQLTranslationKey(List<TranslationKey> translationKeys)
        //public static List<string> GenerateSQLTranslationKey(List<TranslationKey> translationKeys)
        {
            List<string> insertSQL = new List<string>();
            foreach (var translationKey in translationKeys)
            {
                string sqlInsert = String.Format("INSERT #TranslationKey (TranslationKeyID, Tag) VALUES ({0},'{1}')",
                    translationKey.TranslationKeyID.ToString(),
                    translationKey.Tag
                    );
                insertSQL.Add(sqlInsert);
               // sqlInsertTranslationKey.AddStatementLines(sqlInsert);
            }
            return insertSQL;
        }

        public static List<string> GenerateSQLTranslation(List<Translation> translations)
        {
            List<string> insertSQL = new List<string>();
            foreach (var translation in translations)
            {
                string sqlInsert = String.Format("INSERT #Translation (TranslationLanguageID, TranslationKeyID, TranslatedText) VALUES ({0},{1},'{2}')",
                    translation.TranslationLanguageID.ToString(),
                    translation.TranslationKeyID.ToString(),
                    translation.TranslatedText
                    );
                insertSQL.Add(sqlInsert);
            }
            return insertSQL;
        }

    
    }
}



