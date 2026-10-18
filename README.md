# Summary-Excel-Files-in-VBA

Based on my two co-op, I found my work is most excel and calculate either liability of retirement benefit, or the premium of insurance. In addition, both of my work needs to compare the calculated number with last year. Therefore I create this project to make a summary for my work and give employer an easier view.

Assume all the companies or product has their folder and all stored in one main file:


<img src="images/Main%20File.png" width="300">



<img src="images/Subfiles.png" width="600">





The VBA is automatically loop through all the folders, open each file, copy data and paste in a summary file. The tool will also extract the data calculated last year. A summary is generated and highlight the absolute value greater than 30% in red. If the value is 0, the cell will be filled with yellow. If there is a missing value, the cell will be filled with red. 

Furthermore, the VBA will create an issue summary that list 7 types of error. 
1. data is 30% more than last year
2. data is 30% less than last year
3. the value in 2026 is 0
4. the value in 2025 is 0
5. the value in 2026 is missing
6. the value in 2025 is missing
7. the file does not exist in the folder. 

<img src="images/Result.png" width="2000" height = "2000">


## Project Purpose

I designed this project to serve as both a final work summary and a tool for tracking my own progress.

Once the work is completed, the summary gives other employees a clear overview of the results and highlights anything that may need further attention.

I can also run the tool at any point during my work to check what has been completed and what is still missing. This makes it easier to keep track of my progress without having to open and check every Excel file manually.

## Future Applications

Although I built this project based on my actuarial co-op experience, the VBA code is designed as a basic framework that can be adapted to other tasks.

As long as the files are stored in an organized and consistent folder structure, I can modify the code and still generate summaries.

My goal was not just to automate one particular process, but to create something reusable that I could build on in the future.









