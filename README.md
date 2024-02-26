# Solidsport-Analytics-Tool
This Program is creates a powerpoint presentation for customers of the Solidsport Company containing all relevant analytical information of a specified event.

**In order to get the program to work the following dependencies have to be imported/installed:**
1. python 3.12.1 (I think it also works on older versions, I am not sure though)
2. The following packages have to be installed:
   1. pip install selenium
   2. pip install beautifulsoup4
   3. pip install python-pptx
   4. (maybe tkinter, if it isn't preinstalled)
3. **In order to navigate and access the files, as well as to create folders I am using the os tool, thus I don't think the code will work on windows yet** I will try to adapt the code accordingly asap. If you have a Mac this shouldn't raise any errors.
4. **When creating the presentation I am accessing a manually created powerpoint template (simple .pptx file). If this powerpoint doesn't exist on the desktop the program won't be able to create the Powerpoint** I will of course provide the template on here (or send it to you via WhatsApp) _p.s. It makes sense to look at the powerpoint Template, this will give you a good understanding of what Data I am trying to put in the powerpoint_

The **Chronological order** of the code:
1. Import of all relevant packages
2. Definition of all ChromeDriver functions
3. **Formating** and **calculation** functions for the data that will be processed
4. Functions for **accessing** the stored html file according to the specified data type and **extracting** the relevant info
5. **Powerpoint** table functions (I have a few rather complicated Tables that should also be formated neatly)
6. main()
   1. This functions initiates the driver, executes all Chrome functions (clicks the desired buttons), downloads and proceesses the data and finally creates the powerpoint and fills it with the relevant processed Data
8. definition of two GUI functions: retrive_info; gathers all inserted information from the GUI, create_presentation; **this button actaully calls the main()**
9. **configuration of the GUI** and finally the **execution of the GUI**

The **functionalities** of the code:
1. A GUI with Tkinter that is used for userinput
2. A ChromeDriver Bot, logs in on the Solidsport page, goes to the specfied eventpage and then combs through all relevant data
3. A beautifulsoup4 script that saves the html code of all relevant pages and the parses through the html code looking for the relevant data to that html page (e.g. looking for the top five countries viewers watched the stream from by unique views in the geographic html code)
4. A small csv reading spript that analyses the transaction data saved
5. The creation of a powerpoint depending on the template, filled with all processed Data

