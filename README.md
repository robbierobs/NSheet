# powersheetMutilator

This program is built to aid the Enola Diesel Shop dispatch position at inbound. By using python and some libraries, we can use this program to manage the excel sheets that are required for the position, all from one source. This will allow the dispatch position to be much more fluid and more efficient.

## Dependencies

* Python3
* sqlite3
* lxml
* BeautifulSoup4
* getpass
* shutil
* getpass (for more secure password management)

## Installation

1. Download this github repository
2. Install required python libraries using pip
	* sqlite3
	* lxml
	* bs4
	* getpass
	* shutil
3. Execute main program by running `python3 powersheet.py`

## Usage

The bulk of the program is run from the main menu that is displayed when you run the powersheet.py file. The main file includes options to change the powersheets, look at inbound/outbound reports on trains, save the powersheet and webscraping LMIS for relavent information.

### Inbound and Outbound reporting functions

#### Inbound and Outbound Report

The inbound and outbound reporting functions can be used to see what the current state of the powersheet is. This will return the train symbols and their corresponding locomotives to the console. No changes are made with this function.

#### Outbound Trains

This function is a similar report to the "Inbound and Outbound Report" but will only return the outbound section.


### Making changes to the powersheet

As we make changes to the powersheet, the information will store within memory. This information **will not** save until we tell python to save the file. Be sure to save the file so you do not lose your current information after making changes. Each option is accessible through the main menu. After each subsequent function is run, you will be directed back to the main menu to make additional changes.

#### Outbound

We can manage the outbound train assignments per train symbol. Locomotive units are designated by their road numbers with the option to use a direction to designate which way the locomotive is facing (i.e. 9548E would designate that the 9548 is facing East). Addtionally, we can seperate the locomotives from their originating train symbol by using a forward-slash ('/'). This will allow us to show originating train symbols in the "FROM" column of the powersheet (9548E / 4214W would imply that the 9548 and 4214 originated from separate train symbols). As we input the new "build" or "power plan", python will actively search in the inbound side of the powersheet and automatically fill in the "FROM" column on the outbound side. Currently, dates are not implemeneted (18T.28) and will be added later to make the FROM column fully automated.

##### Rob Peter / Pay Paul

This function is use to add additional train symbols that we do not terminate/use on a regular basis. We can use this option to add the train symbol (66X.27), store the unit numbers and apply them to an outbound consist.

##### Search for open engines

Search for open engines is a way for us to see which inbound units have **NOT** been applied to an outbound consist.

### Web Scraping

There is a built in option that allows us to web scrape LMIS for information needed for inbound/outbound reports at the dispatch desk. Some of this information is as follows:
	* Road number
	* Model
	* Group designation
	* Fuel capacity
	* Air flow due date
	* Cab signal due date
	* Maintainence due date
	* EPA due date
	* If labs are due
	* If lubes are due
	* PTC Equipment
	* Energy Management Equipment
	* Locomotive assignment and alternative shop
	* and more
All scraped information **is stored** in a local sqlite3 database that resides in the same folder as powersheet.py. Currently the database isn't actively queried for information but does store information on units. In the future, this will allow to keep detailed information on each locomotive that comes through Enola Diesel. We can track what dates were updated and inbound reporting column have been in development. This will allow the inbound crew to input what work was done, by whom, on which date to allow a more comprehensive system to track information.

**Note** - As the web scraping is done per call, it is normal for the program to throw and error and crash. This will be corrected in the future and will not be an issue at all when the OBDC driver and Teradata server information is utilized in this program. Until then, restart the program and continue to work. It is also common for the HTTP request to time out when web scraping due to the infrastructure of LMIS. Restart the program and continue to work.

### Making work packets

We can create work packet cover sheets from inside the program itself. By inputting your LMIS username and password, you will be prompted to input road numbers. **slugs need to be input with a leader '0' (i.e. 764 would be '0764')**. Once road numbers are input, relavent information will be returned to the console with options to add worksheet headers and commnets. Samples, Lube, Cab Signal, MI date and due, EPA due rows are automatically filled in. Once work troubles are input (or left blank), python will save a copy of the cover sheets (MI_CoverSheets.xlsx & UR_CoverSheets.xlsx). These files can be emailed from an external email address (due to the limitations and restrictions imposed on Python from Norfolk Southern Corp's IT department). I have utilized a script to collect the cover sheets and email them to my internal address via neomutt and a python script. This will be included in the future

### Future

As I continue to the develop the software, changes are bound to happen. I will attempt to keep this as up-to-date as possible. Many features are still planned or currently being developed (including SQL queries from Norfolk Southern's database to allow faster acquisition of information instead of depending on LMIS web scraping).
