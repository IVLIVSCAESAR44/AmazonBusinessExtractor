The program attached is to pull the Business Sales reports from Amazon. This is a fully automated ETL pipeline which transforms and loads the data at the end. Below is an explanation and setup guide.

The program goes into Amazons vendor central terminal and pulls the Business Sales report, thanks to no API this was the only way to do it. Thanks Jeffrey. 

One issue that is worked through on this program is Amazon having 2FA for bot protection. However, them working on user satisfaction caused it to be bypass-able. They have some sort of trust variable programmed into the program so if you successfully do the 2FA about 5 times, it stops asking entirely. After doing it 5 or 6 times, I take the browser data from those sessions and have the Selenium automation instance run off of them. 

After the report is pulled, the rows and columns are formatted to fit the data structure in the database. The rows are then inserted one at a time using a loop that goes off the amount of rows in the dataframe. At the end the file is moved into an archive folder to signal completion.

Setup(Line numbers should be on in your IDE for ease of following):
1. Download Python
2. Using PyPi, pip install selenium, pandas, pyodbc, and shutil. Some of these may already be installed with python default directories.
3. At lines 17 and 18 you'll be putting the location of your user browser data and where you want the file to go once its downloaded respectively. The former is to bypass the 2FA, so if you haven't already, you wanna log in and log out until the 2FA prompts disappear.
4. At line 123, put the path where the file is located after download.
5. At line 147, in the connection, input the server name and database name where prompted. The ODBC driver should be already installed on your computer as it usually comes with windows.
6. At line 160, input the table name to which you want it to load the data into.
7. At line 226, set the file location including the file name in the "Old file name" location. In the File Archive Location, input the location of where you want the used file to be archived, including what you want it to be called.

The program is ready to run.


Created by Adam Nitecki(IVLIVSCAESAR44)

