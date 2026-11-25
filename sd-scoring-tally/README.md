# SD Scoring Tally

Adds up the "Important Things" scores (cells J5:W12) from every SD Individual Scoring `.xlsx` in a folder.

## Usage

```bash
pip3 install openpyxl
python3 tally_scores.py "/path/to/folder"
```

Writes `SD_Scoring_TOTALS.xlsx` into that folder:

- **Totals**: the matrix in its original layout with J5:W12 summed (other sections copied from the first file)
- **Averages**: same layout, each J5:W12 box = total ÷ number of files (blank boxes count as 0)
- **Symbols**: averages as ● strong (≥1.5), ◐ moderate (0.75–1.49), ○ weak/none (<0.75). Limits are at the top of the script
- **Alignment**: Important Things ranked with coverage status, targets with no/weak links, and boxes where people disagreed (0 vs 2)
- **Checks**: scores other than 0/1/2, blanks, boxes not everyone scored, mismatched headings

Note: the template's icon-set conditional formatting isn't carried over; the Symbols sheet uses ● ◐ ○ characters instead.
