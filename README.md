# Stock Market Analysis

[![excel-screenshot]]
<!-- TABLE OF CONTENTS -->
## Table of Contents

* [About the Project](#about-the-project)
  * [Built With](#built-with)
* [Getting Started](#getting-started)
  * [Prerequisites](#prerequisites)
  * [Installation](#installation)
* [Description](#description)
* [Contact](#contact)



<!-- ABOUT THE PROJECT -->
## About The Project

The purpose of this project was to become familiar with the basics of coding using VBA: naming variables; conditionals; for loops; etc. I believe Excel was used because of the program's row/column layout. This provided an early opportunity to get a visual sense of how data can be stored and manipulated, while still being familiar to non-coders. Additionally, this was the first time creating and uploading to a GitHub repository. 

### Built With
VBA

<!-- DESCRIBING PROJECT -->
## Description

For this subroutine, three for loops were implemented within an original for loop that progressed through all the worksheets of the excel file, e.g. data for 2014, 2015 & 2016. The nested for loops accomplished the following, respectively:
<ol>
    <li>Added the name of the stock ticker/symbol (e.g. A) to column I when a row's value no longer matches the previous row's value; calculated the stock volume across the range of cells containing the same ticker and added this to column L. This would continue to accumulate with the <i>Else</i> condition until a new ticker value was detected in the <i>If</i> condition.</li>
    <li>In this for loop, two sets of conditional statements are used, one within the other. The first determines the opening stock value of the year as well as closing, whose difference is used for column J's 'Yearly Change.' The second uses these values to determine column K's 'Percent Change.'</li>
    <li>The final for loop formats the color of the cell in column J. If stock has increased, the cell is changed to green; decreased - red; no change - yellow. </li>
</ol>

<!-- CONTACT -->
## Contact

John Jostes - [https://www.linkedin.com/in/john-jostes-386b841a0/](https://www.linkedin.com/in/john-jostes-386b841a0/)

Project Link: [https://github.com/jjostes/stock-market-analysis](https://github.com/jjostes/stock-market-analysis)


<!-- IMG LINK -->
[excel-screenshot]: https://github.com/jjostes/stock-market-analysis/blob/master/results/2014_stock_data.png
