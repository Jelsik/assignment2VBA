
# Assignment 2: Re-Stocking

## Purpose

The purpose of this project was to take a previously functioning VBA macro and refactor it.
The hypothetical boundaries of the dataset this code analyzed were going to be expanded, and thus the code
should be re-analyzed to ensure that it is optimized to run through larger datasets efficiently.

## Results

### Differences

After refactoring the subject code block, the resulting changes can be visualzed below.

#### Original

![Original Code Block](https://github.com/Jelsik/assignment2VBA/blob/135e25126cefbfc9b8a059a6f4036c7fa170e355/codeoriginal.PNG)

and

#### Refactored

![Edited Code Block](https://github.com/Jelsik/assignment2VBA/blob/135e25126cefbfc9b8a059a6f4036c7fa170e355/coderefactored.PNG)

### Process

The major change in the program was the elimination of an entire for loop. If the data set is to increase in size,
then the code must be reviewed in order to make it more efficient.
	This review was executed by declaring a set of arrays. These arrays could then be used to collect the results
of all of the stock tickers in the same original loop, without having to take the time to process and output the
data using a single individual variable, that had to be output and then be cleared. By collecting all of the data at once
a for loop has been eliminated in it's entirety, and all of the required stock data has been stored, and can posted
to the worksheet in a seperate, smaller environment.

### Outcome

The efficiency of removing a for loop results in a time save of over 500% compared to the original.

#### Original Run Time

![Original Run time](https://github.com/Jelsik/assignment2VBA/blob/135e25126cefbfc9b8a059a6f4036c7fa170e355/Resources/VBA_Challenge_2018Original.PNG)

compared to

#### New Run Time

![New Run Time](https://github.com/Jelsik/assignment2VBA/blob/135e25126cefbfc9b8a059a6f4036c7fa170e355/Resources/VBA_Challenge_2018Refactored.PNG)

### Outcome Continued

Though a few tenths of a second may seem like a minor difference now, if the dataset is significantly lenghtened,
those results will pay out great dividends exponentially.

## Summary

Refactoring code as a process provides great advantages. By applying a second set of eyes to even a completed project,
or just returning to previous work yourself in a new headspace; hidden or lost potential in the macro may be recovered.
Potential disadvantages may simply include the old adage of "if it's not broke, don't break it", wherein  changes
may inadvertently cause negative results.
	Such is not the case in this refactoring of this VBA code, however. Although the original code was functioning
perfectly fine reassigning the same variable after outputting each cell inside it's own loop, better results are made
 by deploying arrays instead in the refactored code. Not only do the arrays make the code neater and more readable, by
defining more variables the code has become more capable of accepting a change in the dataset.
