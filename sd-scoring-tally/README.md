# SD Scoring Tally

Scores are impact scores: 2 = high impact, 1 = some impact, 0 = support.

Adds up the "Important Things" scores (cells J5:W12) from every SD Individual Scoring `.xlsx` in a folder.

## Usage

```bash
pip3 install openpyxl
python3 tally_scores.py "/path/to/folder"
```

Writes `SD_Scoring_TOTALS.xlsx` into that folder:

- **Key Takeaways** (first tab): plain-English summary for sharing - top targets and ties, targets to protect, coverage gaps, targets to refine, disagreements, next steps
- **Totals**: the matrix in its original layout with J5:W12 summed (other sections copied from the first file)
- **Averages**: same layout, each J5:W12 box = total ÷ number of files (blank boxes count as 0)
- **Symbols**: the averages with the template's icon formatting (○ support <0.75, ◐ some impact 0.75–1.49, ● high impact ≥1.5). Adjust the levels in Excel via Conditional Formatting → Manage Rules, or change HIGH_MIN / SOME_MIN at the top of the script (these also drive the Alignment sheet)
- **Alignment**: improvement targets ranked by total contribution to the Important Things, with ties in the top 3 (TOP_N) flagged for discussion and targets needing refinement flagged; targets that are the only high-impact target for an Important Thing marked PROTECT; split scores (0 vs 2) listed; final check that every Important Thing has a high-impact target that everyone scored 2
- **Checks**: scores other than 0/1/2, blanks, boxes not everyone scored, mismatched headings

