SET NOCOUNT ON;

CREATE TABLE #Translation (
	[TranslationLanguageID] [INT] NOT NULL,
	[TranslationKeyID] [INT] NOT NULL,
	[TranslatedText] [NVARCHAR](MAX) NULL,
 CONSTRAINT [PK_Translation_Temp] PRIMARY KEY CLUSTERED ([TranslationKeyID] ASC, [TranslationLanguageID] ASC)
)


MERGE dbo.Translation t
USING #Translation s
ON (s.TranslationLanguageID = t.TranslationLanguageID AND s.TranslationKeyID = t.TranslationKeyID)
WHEN MATCHED AND (ISNULL(t.[TranslatedText],'') <> ISNULL(s.[TranslatedText],''))
	THEN UPDATE SET
		t.TranslatedText = s.TranslatedText
WHEN NOT MATCHED BY TARGET 
	THEN INSERT (TranslationLanguageID, TranslationKeyID, TranslatedText) VALUES (s.TranslationLanguageID, s.TranslationKeyID, s.TranslatedText)
WHEN NOT MATCHED BY SOURCE 
	THEN DELETE
OUTPUT
   $action,
   inserted.*,
   deleted.*;
DROP TABLE #Translation

SET NOCOUNT OFF;
