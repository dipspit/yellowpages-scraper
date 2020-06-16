# Yellow Pages Scraper

forked from [scrapehero/yellowpages-scraper](https://github.com/scrapehero/yellowpages-scraper)

# depends
lxml and openpyxl

# changes
* scrapes keywords page from yellowpages.com instead of using user-inputted single use values
* iterates through the pages of each keyword and scrapes data
* changed the way the excel file was created to use openpyxl
* changed command line arguments for location to user inputted values in script run
* if script is run in the same dir where an output.xlsx exists, script confirms deletion before running

# roadmap
* employ duplicate skipping