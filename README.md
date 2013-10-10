GoogleApps SumByColor
=====================

This is a [Google Spreadsheets script](https://developers.google.com/apps-script/) that allows summing up cells depending on their foreground/background color.


Installing the script will make the following functions available in the spreadsheet:

* `getBackgroundColor(<cell specification>)`
* `getForegroundColor(<cell specification>)`

* `sumWhereBackgroundColorIs(<color>, <range specification>)`
* `sumWhereForegroundColorIs(<color>, <range specification>)`
* `sumWhereBackgroundColorIsNot(<color>, <range specification>)`
* `sumWhereForegroundColorIsNot(<color>, <range specification>)`


Please note that `<range specification>` and `<cell specification>` are expressed in **A1** notation, and **must be enclosed in quotes**.

For example, to compute the sum of all the cells in the range B2:F13 that have the background color set to *white*, you should enter the following formula:

    =sumWhereBackgroundColorIs("white", "B2:F13")

Some cells may not have the background set to a color such as 'white', 'gray', but a RGB color like `#6fa8dc`. You cannot guess what the color is, so if you want to find out the color for a cell (for example, `B9`), you should enter this formula in a cell:

    =getBackgroundColor("B9")

and afterwards use this value as a parameter to the two functions above.
