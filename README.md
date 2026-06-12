# call-pm-data
process call pm data like cncs, csbc, mediation

## process-call-data.py
<li>Dynamically figure out which sheet to process. <b>If the sheet name changes, it is possible this script has to change.</b>
<li>Dynamically figure out the file name to write to base on the value on the 2nd column.
<li>Output the max number on the day.

```bash
$ python3 process-call-data.py liu/CCFX_NUM_ACR-700124-2026_06_03-04_03_00__9.xlsx
Processed sheet: Data for NUM_ACR ()
Output written to: vccf.csv
```

<b> Future improvement </b>
<li>Read the previous saved file (like old vccf.csv), and append the new data to the existing file.
<li>Enhanced Error handling on the input xlsx.
