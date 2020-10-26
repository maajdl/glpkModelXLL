# glpkModelXLL
Linear programming within an Excel workbook using ```gmpl-glpk```.  

## Purpose
The ```glpkModelXLL``` addin is a simple way of using ```gmpl-glpk``` from within Excel.   
The ```model``` worksheet contains a gmpl model and information about where the data are stored.   
In this way large models are easily handled using Excel names for the data.   
The model may contain these additional statements: 
* scenario: to specify a scenario name  
* variables: to writes all the variables in a specified table   
* constraints: to write all the constraints in a specified table   

The ```glpkModelXLL``` addin also offers a scenario tool to compare different models.   

## Content
* A Visual Studio C# solution (glpkModelXLL.sln) and source code.
* Three sample model workbooks (Stigler's 1939 diet problem v1.xls, ...)  
* One sample scenarios workbook (scenarios.xlsx)

## Installing and getting started
* copy the whole project at a suitable place on your computer   
* locate the folder    ```/glpkModelXLL/glpkModelXLL/bin/Release/```   
* open the addin  ```glpkModelXLL-AddIn.xll```  with Excel
* open a sample file from the root folder  ```/glpkModelXLL```

## Testing
The  glpkModelXLL  addin adds a ribbon to Excel with buttons that you can try:

| button        | purpose           :| 
| ------------- |:-------------|
| solve         | to solve the gmpl model written in the worksheet  "model" |
| refresh       | to refresh the workbook, specially the PivotTables      |
| mod           | to view the mod file (gmpl model) created by solve      |
| dat           | to view the dat file created by solve, based on the data in the workbook      |
| lp            | to view an lp translation of the model      |
| automatic solve   | if checked, solve will occur automatically for any change in the workbook      |
| automatic refresh | if checked, refresh will automatically follow any solve      |
| scenarios refesh  | will refesh the scenarios or create the scenarios logging sheet      |
  
## References ## 

The ```glpkModelXLL``` addin is build using Excel DNA, see: https://excel-dna.net/   
For the gmpl modeling language and the glpk solver, see: https://www.gnu.org/software/glpk/   
Other interfaces from Excel to glpk, see: https://en.wikibooks.org/wiki/GLPK/Windows_IDEs#Microsoft_Excel_integration

## Screenshot ##

![alt text](https://github.com/maajdl/glpkModelXLL/blob/master/screenshot.png "glpkModelXLL screenshot")

