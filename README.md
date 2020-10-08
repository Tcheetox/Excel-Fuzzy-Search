# Excel Fuzzy Search

Implementation of a fuzzy search algorithm using **RegEx**. The scoring mechanic is inspired by an article of [James Padolsey](https://j11y.io/javascript/fuzzy-scoring-regex-mayhem/).

# Get started

If you wish to try the search with your own data:
1. Override the table in the **DB** sheet and ensure the matrix **Database** fits to your new range. This can be done in Excel via the Name Manager.
2. Close and reopen the file or alternatively use the **Create cache** button in the **Help** sheet.

## How it works

The database is read in memory and stored per line. A new search corresponds to a generated RegEx pattern tested against each of those lines.
Groups and sub-groups are scored **arbitrarily**. The 10 greatest score are displayed in the `SearchPane` by decreasing order. 
Score calculations can be customized in `ExtractCandidateInfo` from the `clsMatchResults` class. The same Sub also generates necessary information for the syntax highlighting.
