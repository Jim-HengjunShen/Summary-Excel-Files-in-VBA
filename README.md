# Summary-Excel-Files-in-VBA

Based on my two co-op, I found my work is most excel and needs to calculate either liability of retirement benefit, or the premium o father insurance. In addition, the work needs to compare the number with last year. Therefore I create this project to make a summary for my work and give employer an easier view of all companies.

Assume all the companies or product has their folder and all stored in one main file:


<img src="images/Main%20File.png" width="300">



<img src="images/Subfiles.png" width="600">





The VBA is automatically loop through all the folders, open each file, copy data and paste in a summary file. The tool will not only contain the data we calculated for this year, but also the data last year. Then it will compare them and highlight the absolute value greater than 30% in red. If the value is 0, the cell will be filled with yellow. If there is a missing value, the cell will be filled with red. 

Furthermore, the VBA will create an issue summary that list 7 types of error. 
1. data is 30% more than last year
2. data is 30% less than last year
3. the value in 2026 is 0
4. the value in 2025 is 0
5. the value in 2026 is missing
6. the value in 2025 is missing
7. the file does not exist in the folder. 

<img src="images/Result.png" width="2000" height = "2000">












