-- Text of each artifact's latest entry file, filled the first time a search
-- needs it, so ?q= can look inside artifacts.

CREATE TABLE search_text (
  artifact_id TEXT PRIMARY KEY REFERENCES artifacts (id) ON DELETE CASCADE,
  n INTEGER NOT NULL,                  -- version the text was read from
  text TEXT NOT NULL
);
