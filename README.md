# How to read/write Excel .xlsx file using Apache POI in Katalon Studio

This is a small Katalon Studio for demonstration purpose. You can download the zip of this project form the [Release]() page, unzip it, open it with your local Katalon Studio.

I developed this project using Katalon Studio v8.2.0 on Mac. This project is not version specific, not platform dependent. It should work in every version of Katalon Studio, on Windows as well.

This project provides a working Test Case script which 
1. reads an Excel file `./Employee.xlsx`, print the data in the sheet into the console
2. updates the sheet with additional rows of data, write the workbook into another Excel file `./Employee.out.xlsx`

The source is here:
- [Test Cases/XLSXReaderWriter](Scripts/XLSXReaderWriter/Script1642118878582.groovy)


The input Excel sheet looks like: ![input](docs/images/input.png)

The output Excel sheet looks like: ![output](docs/images/output.png)

When I ran the Test Case I saw the following messages in the console:

```

```

## 

The Test Case script uses the [Apache POI](https://poi.apache.org/) directly. This script aims to demonstrate a simplest example how to use POI in Katalon Studio. There are a lot of resources about the POI library. So you should search and find articles for you.

The latest verion of Apache POI is v5.0.x, but Katalon Studio v8.2.0 bundles fairly old version of Apache POI: 3.17. You should note this version difference and, mind not to be confused.

You would need to bookmark the [Javadoc of POI v3.17](https://poi.apache.org/apidocs/3.17/)

## 

The source code is originally published at the following site by vektorwebsolutions.com
- https://vektorwebsolutions.com/how-to-read-write-xlsx-file-in-java-apach-poi-example/

The original was written in Java. So I convert it into Katalon Studio's Test Case script.

I modified the code just slightly as I found a few issues to be fixed.
