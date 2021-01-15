# StockHistory
This program looks at the history of components from areas SR03-15 and C00-11. It scrapes the relevant information and puts them onto an excel sheet.

Versions of python and installed modules: 
- python 3.7.8
- selenium 3.141.0
- ChromeDriver 80.0.3987.106
- Visual Studio 16.8.4
- openpyxl 3.0.6
- numpy 1.19.5

Requirements:
Plex login for Wanco

Update 11/23/2020: I'm waiting on test Plex credentials so I can take my own out.

Update 12/4/2020: I now have generic Plex credentials.

Update 12/18/2020: Back to my credentials since the generic account got nerfed in terms of what pages it can access.

In an effort to optimize the storage room and its containers, this program will find all inactive/dead containers in locations SR03 through SR15 and 
C00 through C11. A container is deemed "dead" if there are more instances of "Cycle Complete" and "Container Move" than "Split Container" in the 
"Last Action" column of its history.

Preliminary procedure: The program utilizes the selenium module and the ChromeDriver to create a chrome web driver. This driver opens a new window for Plex and
logs in using the provided credentials.

There's only one method for finding dead containers, but it gets applied to every location. The program starts with location SR03 and goes through each SR location
until, and including, SR15. Next, it moves onto C00 and beyond. The only other part of the search criteria that gets altered is the time frame. The "End Date" field
will be set to one year ago and the "Begin Date" field will be set to an early enough time period that the program can see a container's initial history. For each
container, the program will keep track of two amounts: act and inact. act increases the more "Split Container"'s are found on the "Last Action" column. The only 
instances of "Split Container" that don't count are the ones where the location of the action is "SR RECEIV" because this split is actually the breaking down of
cotainer size for storage purposes. inact increases the more "Cycle Complete"'s and "Container Move"'s are found. If inact is larger than act, the container is 
considered dead. Nothing else happens with active containers, but dead ones are put onto an excel sheet. Entries will have a container number, part number,
location, and a ratio of inact to act to help the storage team figure out which dead containers to prioritize.
The ratio can be on of three options, greater than zero, zero, and infinity. Greater than zero is achieved when act and inact are both larger than zero.
Zero is achieved on the rare case that act and inact are both zero. In addition, containers with a 50/50 split between act and inact are included in the
excel sheet. Infinity is the most common case for a ratio; this occurs when a container's inact is greater than zero, but its act is zero. Since a 
division by zero would crash the program, infinity will be assigned instead.

Given the time frame, I'd expect this program to be ran once a year or so. 
That's a good thing considering how just six locations makes this program run for roughly two hours.
