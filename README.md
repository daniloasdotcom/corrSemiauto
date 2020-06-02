# corrSemiauto

A python script to automate the production of R Markdown scripts for Pearson's Correlation reporting

## Repository files

1. corr.py
2. Estimating.xlsx
3. example.xlsx
4. StepByStep.pptx

## Step By Step

1. Edit the corr.py file in the locations indicated in the script comments
2. Run the corr.py file and a .txt file will be generated with the necessary information for the wheels in Rstudio
3. In Rstudio:
	1. Open a new R Markdown file
	2. Erase everything that was generated
	3. Paste the text from the .txt file
	4. Run

**At the end**

1. A .docx file with
	1. Your data list
	2. A correlation matrix in numbers
	3. A graphical correction matrix
	4. And the significance of each of the tested correlations
1. An .xlsx file
	1. With the correlation table generated
