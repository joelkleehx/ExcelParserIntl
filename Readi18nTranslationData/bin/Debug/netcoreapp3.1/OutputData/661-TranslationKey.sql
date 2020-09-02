SET NOCOUNT ON;

CREATE TABLE #TranslationKey (
	[TranslationKeyID] [INT] NOT NULL,
	[Tag] [NVARCHAR](255) NOT NULL,
 CONSTRAINT [PK_TranslationKey_Temp] PRIMARY KEY CLUSTERED ([TranslationKeyID] ASC)
 )

SET IDENTITY_INSERT TranslationKey ON

MERGE dbo.TranslationKey t
USING #TranslationKey s
ON (s.TranslationKeyID = t.TranslationKeyID)
WHEN MATCHED AND (t.[Tag] <> s.[Tag])
	THEN UPDATE SET
		t.Tag = s.Tag
WHEN NOT MATCHED BY TARGET 
	THEN INSERT (TranslationKeyID, Tag) VALUES (s.TranslationKeyID, s.Tag)
WHEN NOT MATCHED BY SOURCE 
	THEN DELETE
OUTPUT
   $action,
   inserted.*,
   deleted.*;

SET IDENTITY_INSERT TranslationKey OFF

DROP TABLE #TranslationKey

SET NOCOUNT OFF;
