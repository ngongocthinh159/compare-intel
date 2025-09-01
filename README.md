## Inter compare
### K1 compare
```
python row_compare_offset.py --manual-path ".\inter-manual.xlsx" --manual-sheet "K1" --auto-path ".\inter-auto.xlsx" --auto-sheet "K1Intermediate" --num-cols 10 --manual-start 8 --auto-start 2 --offset 0
```

### K2 compare
```
python row_compare_offset.py --manual-path ".\inter-manual.xlsx" --manual-sheet "K2" --auto-path ".\inter-auto.xlsx" --auto-sheet "K2Intermediate" --num-cols 7 --manual-start 8 --auto-start 2 --offset 0
```
----------------
## Final report compare
### Command
```
--manual-path ".\report-manual.xlsx" --auto-path ".\report-auto.xlsx" --sheets "Summary" "M1 IC" --rows 60 --cols 25 --start-row 1 --start-col 1
```
--sheets "Summary" "M1 IC" ... : can be mutiple sheet names

--rows 60 : number of rows to compare

--cols 25 : number of cols to compare

--start-row 1 : starting row to compare (--start-row + --rows - 1 = last compare row)

--start-col 1 : starting col to compare (--start-col + --cols - 1 = last compare col)
