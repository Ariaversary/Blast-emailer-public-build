# Blast-emailer-public-build

# This program's intended use is for Microsoft Azure, you may edit the code to fit your use as needed (including Client, Tenant and Client Secrets)

# Disclaimer: Use at your own risk!

How to use the blast emailing software:


Step 1:

Enter Client Secret and ID, and Tenant ID using Microsoft Azure

Change the [SENDER EMAIL] line inside the config.json.

![image](https://github.com/user-attachments/assets/f35631ce-fc83-40e4-90e8-9edd083b2b7b)

**You can use notepad to access the code and change it there as labelled**



Step 2:

Download the .msg file from Outlook to save the email template as required.

If you want to edit the email's template, please send an email to a random account with the template you want, then download that .msg file instead to be used

**Please have the subject of the email be the subject you want to be sent out as well**



Step 3:

Run the application, then follow the steps required to send out the emails.

Follow the steps required:

	a) You would be required to select the .msg file as your email template
		![image](https://github.com/user-attachments/assets/012575f7-9e09-45e4-992c-0b8c46a7e04b)

	b) You would then need to select your excel file as a database. (Make sure it has an email column in the sheet you want to use.)

	c) Select the images that were used in the original email, from *TOP* to *BOTTOM*

		Example: If you had a sales image in the middle of the email, then have a business card at the bottom, Please select the sales image first, then business card image next.

  		**You can also enter hyperlinks now if you'd wish on each individual image.**

    		![image](https://github.com/user-attachments/assets/823527c4-080e-4b67-8390-86834e810564)


	d) Enter the name of the specific sheet you want to use. (Make sure the spelling is exact)

			Example: Sheet name "Malaysian Government (135)"

				Make sure you enter the name 1 to 1.

    			![image](https://github.com/user-attachments/assets/d86cf29f-d78b-485f-8b6a-5e05695d2d89)
	
			In this case, you would need to enter exactly "OCW5CFF" with exact spacing and text for it to work
   

	e) Enter the sending delay between each emails in seconds (Recommended is 15 seconds, to prevent auto-spam filters)

		You can also enter emails you would want to CC during this part of the program

    		![image](https://github.com/user-attachments/assets/a801c614-c6c3-4a72-85b7-f171fd05327c)



If everything was selected properly, the program would then start sending all the emails to every row in the selected excel sheet.

There would be a window popped up showing the sending process.

![image](https://github.com/user-attachments/assets/d51bd8ef-8d56-4647-a655-925421edd493)



**Do note that if the email column exists, but the row being sent does not have an email or multiple emails in the row, there would be an error, but it wont stop the sending process**
