# WebScraper


Web scraper that collects data from a single page importing it into excel file, written with Python.  




##### Known issues

* Occasionally the script will return "No Level 1 Projects" message when Level 1 Projects present. 
Running the script second time fixes the problem. 



#### **_Version 2_**

* Replaced HttpNtlmAuth library with Selenium
* URL is now hard-coded
* The 2 types of projects are now combined into a single list and then proccessed, reducing duplicated code and overall code length.
* The code is moved into one main() function 
* Removed unused libraries
* Console output has been formatted for better readability 
* Openpyxl UserWarning (Unknown extension is not supported and will be removed) has been ignored (all warnings ignored) and thus is not displayed in the console.
* Added new and edited existing comments
* Added icon
