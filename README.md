# VBA_challenge
VBA_ challenge. Module 2, Assignment 2

Please see the attached Excel document and VBA script document for additional information. 
Screenshots can be found in the Images folder. (to make it easier to see, I only captured the first 20 rows for each year, to see more of the data please view the Excel document).


## Multiple Year Stock Data

I created a script that loops through all the stocks for one year and outputs the following information:

-The ticker symbol

-Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

-Highlights positive change in green and negative change in red for the yearly change column.

-The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

-The total stock volume of the stock. 

I then added headers(column names) for each of these new columns.

![part 1](https://user-images.githubusercontent.com/120147552/210837701-63289e13-ed97-4b86-b3d6-58635cd27424.png)

Then I added functionality to my scrip to return the stocks with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 
I added column and row names to make this information easier to identify.

![part 1](https://user-images.githubusercontent.com/120147552/210837717-04a0de07-ac1b-4714-a104-6212412830d9.png)

I then created a loop that would run the above loops on every worksheet (that is, every year) at once; as seen below:

#### 2018

![2018](https://github.com/BrendaWardhaugh/VBA_challenge/blob/main/Images/2018.png)

#### 2019

![2019](https://github.com/BrendaWardhaugh/VBA_challenge/blob/main/Images/2019.png)

#### 2020

![2020](https://github.com/BrendaWardhaugh/VBA_challenge/blob/main/Images/2020.png)


## VBA Script
#### Please see VBA Script file attached for clairification and to better see the order the script/loops appear in.

For the first loop (to get the ticker symbol, yearly change, percentage change and total volume): 

I labled my "Headers" and formated the cells accordingly:

![script1](https://user-images.githubusercontent.com/120147552/210845003-ceec3741-4705-436f-8066-2c8b0638791c.png)

![script5](https://user-images.githubusercontent.com/120147552/210845691-a2ec4511-e326-42f2-851a-ad805086b1ef.png)

I identified my variables. 

![script 1](https://user-images.githubusercontent.com/120147552/210844205-6ff20191-fbdb-4b93-ab86-45017b3a7a8c.png)

Then ran the first loop:

![script 2](https://user-images.githubusercontent.com/120147552/210844701-d88619f6-8eaf-4dc3-8224-37f63b95e222.png)

![script7](https://user-images.githubusercontent.com/120147552/210846384-308c0852-7394-43a2-abaf-85c436acb2d9.png)

Next I identified my variables and looped through to get the return the stocks with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume":

![script8](https://user-images.githubusercontent.com/120147552/210846845-73dd7181-6773-4b00-bf42-dc812f6bc210.png)

![script3](https://user-images.githubusercontent.com/120147552/210845229-5a16b38a-9e59-4aa6-bcc3-67a6219ed11f.png)

And placed them in the appropiate location in the sheet with correct format:

![script4](https://user-images.githubusercontent.com/120147552/210845515-a8472897-1d6d-4ab6-8080-a63f16082c34.png)

Next I colour coded the Yearly Change column according to positive change in green and negative change in red:

![script6](https://user-images.githubusercontent.com/120147552/210846234-870dc1b7-0d3c-4ce2-b5c4-4bc434aeec61.png)

Finally, I made it so that it would loop through all of the worksheets by placing my other loops inside of this loop:

![script9](https://user-images.githubusercontent.com/120147552/210847244-edf0861a-0cb0-4e63-a7c0-a9a86f253f2a.png)

