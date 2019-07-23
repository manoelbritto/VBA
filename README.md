# VBA

# Project: Stock market data

This project has the goal to put in practice my VBA skills. For this reason, I gathered a dataset from a stock market. And from this point, I developed three macros to iterate within all worksheets. Moreover, the results you can check below (Result header).

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

You should have your development view activated in your Microsoft Excel application. You will know if you have a menu called Developer. In case you don’t have, then try to search “ Activate the Developer Tab in Excel”

## Running this development

At the Developer Tab, you can search for a button called Macros, then you will find these three macros:
calculate_Vol_easy
calculate_Moderate
calculate_Hard

### Result

Macro calculate_Vol_easy
It will loop through one year of stock data for each run and return the total volume each stock (Column J) had over that year.
And display the ticker symbol (Column I) to coincide with the total stock volume.

Macro calculate_Moderate
The ticker symbol (Column I).
Yearly change from opening price at the beginning of a given year to the closing price at the end of that year (Column J).
The percent change from opening price at the beginning of a given year to the closing price at the end of that year (Column K).
The total stock volume of the stock (Column L).

Macro calculate_Hard
This includes everything from the macro calculate_Moderate.
Also, it returns the stock with the "Greatest % increase" (Column N), "Greatest % Decrease" (Column O) and "Greatest total volume" (Column P).

## Built With

Excel VBA

## Versioning

Version 1.0

## Acknowledgments

1. VBA script
2. Functions
3. Conditions
4. Loops
5. Code Reusability
