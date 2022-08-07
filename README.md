# VBA Of WallStreet

## Overview of Project

### Purpose
Steve wants to analyze stocks within the last couple of years to see what are good stocks to invest in. The purpose of this project is to develop a script where it will provide the total volume of stocks based on its ticker and the return rate within that year. Additonally, The script will refactored to be able to handle multiple years worth of data. 

## Results

### Before Refactoring
Before refactoring the VBA script, When running the script it would take many seconds to output data for the analysis.


![2017 Before Results](https://github.com/40super/stock-analysis/blob/main/resources/Slow%202017.png?raw=true)
![2018 Before](https://github.com/40super/stock-analysis/blob/main/resources/Slow%202018.png?raw=true)

It took around 16 to 17 seconds to analyze the data and output the results onto the workbook. This is very long and can be a issue when facing larger dataset.

### After Refactoring
One of the issues noticed with the script prior to refactoring involves the nested for loops. Most of the computation occurs within the nested loop which results in slower times to analyze. The way to fix this is to do less computation within the for loop in which this case was to use arrays for all variable used withing the loop. As a result, the time for the script to run had drop significantly.


![2017 After Results](https://github.com/40super/stock-analysis/blob/main/resources/Timer%202017.png?raw=true)
![2018 After Results](https://github.com/40super/stock-analysis/blob/main/resources/Timer%202018.png?raw=true)

After refactoring, the code time when from an average of 16.5 seconds to 0.3 seconds. This is significant because when handling massive amount of data, the run time can increase massively which can result in alot of loss time.
### 2017 VS 2018
Looking at both years, 2017 had the most biggest return across most stocks while 2018 had the worst return where most stocks were in the negatives.


![2017](https://github.com/40super/stock-analysis/blob/main/resources/2017%20Analysis.png?raw=true)

![2018](https://github.com/40super/stock-analysis/blob/main/resources/2018%20Analysis.png?raw=true)


Furthermore, One key information that would be useful to Steve is that within those two years, ENPH and RUN are the only two stocks that had a positive return. Knowing this, Steve can invest in these two stocks having a higher chance of getting a return in the coming years.

## Summary

### Refactoring in General
Refactoring is a skill that can create many opportunity to provide clean and efficient code which can result in better output. There are many advantages to refactoring which can help make code more versatile on a wide range of data and also create cleaner code. Some disadvantage to refactoring is that it take time which can be limited if a project has a deadline. Additionally, Refactoring poorly can cause more issues and even make time to run slower.

### Refactoring in VBA
Refactoring the original script had improve runtimes significantly which is a major advantage but the only disadvantage that I think is the time it take to look through the code and determine what has to be refactored for run times to be better.
