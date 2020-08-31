using System;
using System.Collections.Generic;
using System.Text;

namespace Readi18nTranslationData
{
    public class SQLStatement
    {

        public List<string> StatementLines { get; set; }

        public void AddStatementLines(string insertStatement)
        {
            StatementLines.Add(insertStatement);
        }

        public SQLStatement()
        {

        }
    }
       
}
