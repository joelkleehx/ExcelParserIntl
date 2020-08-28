using System;
using System.Collections.Generic;
using System.Text;

namespace Readi18nTranslationData
{
    class Translation
    {
        public int TranslationID { get; set; }
        public int TranslationLanguageID { get; set; }
        public int TranslationKeyID { get; set; }
        public string TranslatedText { get; set; }
    }
}
