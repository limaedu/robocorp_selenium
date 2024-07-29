<h1 align = 'center'>  Robocorp Selenium</h1>

This repository uses Selenium integrated with Robocorp to extract information from the website https://www.aljazeera.com/ and outputs it to an Excel file. To run this from Robocorp, start the main task with the following work items:

- search_phrase: The category to be searched on the site.
- month_filter: An integer used for date filtering with the following logic:
    - 0 or 1: only the current month,
    - 2: current and previous month,
    - 3: current and two previous months, and so on.