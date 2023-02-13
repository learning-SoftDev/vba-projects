# vba-projects

<h3>Automation App 1 </h3>
<img src="https://user-images.githubusercontent.com/96763817/218365981-fe061a35-af5b-4601-88dc-eb320f73da25.png" alt="img_5" width="500px">
<img src="https://user-images.githubusercontent.com/96763817/218365974-ca7c7b64-0de0-47c7-a204-3ea12520b7ac.png" alt="img_2" width="500px">
<img src="https://user-images.githubusercontent.com/96763817/218365970-a90c9492-0a84-4bea-a134-a4ac4bd2927c.png" alt="img_1" width="500px">
The main gist of the app is to update a certain excel input called "Master Tracker". Basically, the app pulls in data from local excel inputs and does the following: <br><br>

- Updates details of the master tracker based from the selected input. The master tracker contains the primary key for joining.<br>
- Performs an XlookUp on master tracker and more than 3 files simultaneously with a click of a button.<br>
- Automatically highlights rows that requires the attention of the admin.<br>
- Shows a progress bar so that the user will have an idea on the progress of the app in the background.<br>
- Have proper validation to avoid bugs i.e. checking of required columns before the app runs the main process it needs to run and checking of error cells.<br>
- Only displays required files based from different criterias to avoid the generation of unnecessary files to avoid waste of time.<br>


<h3>Automation App 2 </h3>
<img src="https://user-images.githubusercontent.com/96763817/218365977-9bc3de27-5f7e-4e5d-a8b6-95a285b11b44.png" alt="img_3" width="500px">
This app takes in a table from sharepoint list then performs data transformations and cleansing with a click of a button. It then generates an excel file which will be used on a different application. The main goal of the app is for the users to have a finished product ready for usage after clicking the Run button. Specifically, it: <br> <br>

- Removes non-alphanumeric characters<br>
- Informs user which cells contains an incorrect email format<br>
- Performs a checking on the contain of certain columns and informs user if an invalid content was noted<br>
- Since the column names of the sharepoint list and the other application differs, the app offers a tab wherein the users can map the columns from these two different datasources and then follows the naming convention of the columns based from the other app. This way, the generated file will be ready for usage.
- Formats the numbers and dates automatically to avoid errors on the usage for the other app.<br> 


<h3>Automation App 3 </h3>
<img src="https://user-images.githubusercontent.com/96763817/218365980-adf61d14-07f3-45c8-b286-f54a5bd4e089.png" alt="img_4" width="500px">
This app focuses on validation of inputs. Every busy season, the team encounters difficulties on incorrect inputs by the requestors of the service. Hence, I thought of an initiative so that we can avoid these circumstances. <br>
The old file only contains the column headers and instructions on how to fill out the file (which is prone to mistakes from the requestors side). The new app that I developed and proposed to the team removed a huge percentage of the risk of this kind of thing from happenning. <br><br>
Basically, the app: <br>

- Creates a data validation based from the allowed values specified in the other tab of the app. <br>
- Copy pasting of contents in excel also removes the data validation, this is one of the reason why the old file is prone to mistakes since the requestor sometimes copy paste’s old details hence will be incorrect. The developed app will automatically delete the content and retains the data validation if ever the user will copy and paste’s incorrect inputs. <br>
- After the user fills out a single cell, it will highlight the required columns to be filled out. If ever the user leaves a required column as blank, he/she will not be able to save the file.<br>
- Validates columns which needs email contents if the provided input has a valid email address format.<br>
- Since some columns’ requirement depends on the contents of other columns, the app also scans the file if it needs to tag the said column as a required (yellow highlight) or not.<br>


