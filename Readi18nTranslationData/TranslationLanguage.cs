using System;
using System.Collections.Generic;
using System.Text;

namespace Readi18nTranslationData
{

    public enum Language
    {
        USEnglish = 1, 
        Finnish = 2,
        French = 3,
        Swedish = 4,
        Norwegian = 5
    }

    class TranslationLanguage
    {
        public int TranslationLanguageID { get; set; }
        public string Name { get; set; }
        public string Culture { get; set; }
        public bool IsDefault { get; set; }
    }
}
