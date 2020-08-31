using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace Readi18nTranslationData

{



   

    class Program
    {

       
      
        static void Main(string[] args)
        {


            List<string> completeSQL = new List<string>(); 

             String fileName = @"C:\Users\JCK0412\source\repos\ExcelParser\Readi18nTranslationData\data\HQ-Internationalization-v2-Finnish.xlsx";
            //String fileName = @"C:\Users\JCK0412\source\repos\ExcelParser\Readi18nTranslationData\data\sample.xlsx";


            //SQLStatement sqlInsertTranslationKey = new SQLStatement();
            //sqlInsertTranslationKey.StatementLines = new List<string>();

            //sqlInsertTranslationKey.AddStatementLines("fred");
            //sqlInsertTranslationKey.AddStatementLines("fred");

            //sqlStatement.StatementLines.Add("fred");
            //sqlStatement.StatementLines.Add("bob");
            //sqlStatement.StatementLines.Add("jim");





            // use Tim Corey's course to set up .Net Core config files: https://www.iamtimcorey.com/courses/515189/lectures/9448253





            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    // WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    List<WorksheetPart> worksheetParts = workbookPart.WorksheetParts.ToList(); 

                    foreach (WorksheetPart worksheetPart in worksheetParts)
                    {
                        ProcessWorksheet(worksheetPart, sstpart, sst);
                       
                    }
                }
                Console.ReadLine();
            }
        }

        public static void ProcessWorksheet(WorksheetPart worksheetPart, SharedStringTablePart sstpart, SharedStringTable sst)
        {
            Worksheet sheet = worksheetPart.Worksheet;

           

            var rows = sheet.Descendants<Row>();
           
            bool loadData = false;
            char englishTextColumn = 'C';
            char tagColumn = 'E';
            List<Translation> translations = new List<Translation>();
            List<TranslationKey> translationKeys = new List<TranslationKey>();
            int translationKeyID = 0;
            int translationID = 0;

            var counter = 0;

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
                               Console.WriteLine(englishText);
                                counter++;

                            }

                            if (columnLetter == tagColumn)
                            {
                                tag = cellValue;
                                Console.WriteLine(tag);
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
            Console.WriteLine(counter);
            Console.ReadLine();
        //    LoadTranslationKey(translationKeys);
            LoadTranslation(translations);

            
        }

        //public static List<string> GenerateSQLTranslationKey(List<TranslationKey> translationKeys, SQLStatement sqlInsertTranslationKey)
        public static List<string> GenerateSQLTranslationKey(List<TranslationKey> translationKeys)
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

        public static void LoadTranslationKey(List<TranslationKey> translationKeys)
        {
            string[] initialSQL = File.ReadAllLines(String.Format(@"{0}\sqlTemplates\TranslationKeyInit.txt", Environment.CurrentDirectory));
            List<string> insertSQL = Program.GenerateSQLTranslationKey(translationKeys);
            string[] loadSQL = File.ReadAllLines(String.Format(@"{0}\sqlTemplates\TranslationKeyLoadRecords.txt", Environment.CurrentDirectory));

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\temp\661-TranslationKey.sql"))
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
    }
}



