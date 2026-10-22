The archive latency script runs on Python's integrated development environment, Spyder.

****PLEASE ENSURE YOU ADD YOUR EARTHDATA AND CDSE CREDENTIALS TO THE 'earthdata_credentials','earthdata_token' and 'cdse_credentials' text files.****

**NASA EARTHDATA - CREATE AN EARTHDATA LOGIN**
Earthdata Login provides access to NASA Earth science data services.
Step 1: Open:
https://urs.earthdata.nasa.gov/
Step 2: Click “Register”.
Step 3: Complete the required fields, choose a username and create a password. Keep a record of your username: do not assume that it is your email address.
Step 4: Submit the registration form.
Step 5: Open the activation email from Earthdata and click the activation link. Your account must be activated before you can sign in.
Step 6: Return to the Earthdata Login page and sign in to confirm that the account works.
Credentials to keep privately:
•	Earthdata username
•	Earthdata password
Step 7: Save your Earthdata username and password in the ‘earthdata_credentials’ text file. Replace ‘USERNAME_HERE’ with your username, and ‘PASSWORD_HERE’ with your password.

Step 8: After signing in, open “Generate Token” in your profile and follow the instructions. Use “Show Token” to view and copy it into the ‘earthdata_token.
(Please note: Earthdata user tokens work with compatible applications and currently expire after 60 days. Check the expiry date and replace the token when necessary.)

**COPERNICUS DATA SPACE ECOSYSTEM (CDSE) - CREATE AN ACCOUNT**
Step 1: Open:
https://dataspace.copernicus.eu/
Step 2: Click the 'login’ icon near the top-right corner, select “Your profile” tab, then select “Register”.
Step 3: Complete the required information, enter your email address and create a password.
Step 4: Read and accept the required terms and privacy policy, then submit the registration form.
Step 5: Open the verification email and click “Verify email address”. Check your spam or junk folder if necessary.
Step 6: Return to the CDSE website and sign in using your registered email address and password.
Credentials to keep privately:
•	CDSE email address
•	CDSE password
Step 7: Save your CDSE username and password in the ‘cdse_credentials’ text file. Replace ‘USERNAME_HERE’ with your username, and ‘PASSWORD_HERE’ with your password.
Registration guide:
https://documentation.dataspace.copernicus.eu/Registration.html

**BEFORE RUNNING THE SCRIPT**
Confirm that:
•	Your Earthdata account is activated, and you can sign in.
•	Your CDSE email is verified, and you can sign in.
•	You have entered and saved the required credentials and token in the appropriate text files.
Keep passwords, API keys and tokens private. Do not upload them to GitHub or include them in shared results, screenshots or error reports. Any local files containing credentials must be excluded from Git tracking.

__________________________________________________________________________________________________________________________
If you do not have Spyder, please download and install it from:
1. https://www.anaconda.com/download
OR
2. https://www.spyder-ide.org/download


********************************STEP ON HOW TO RUN THE SCRIPT**********************************

NOTE: Extract the zip file and open the script from the extracted folder.
*** PLEASE DO NOT USE VPN WHEN RUNNING THE SCRIPT to avoid miscapturing of certain latency metrics.**

PLEASE GO TO STEP 1.
____________________________________________________________________________________________________________________
A few packages might need to be installed for the script to run effectively. To check this, perform Step 0.

Step 0. Install the following packages by pasting the line below into the Spyder console (right bottom quadrant) and pressing the 'Enter' button:

pip install requests pandas matplotlib speedtest-cli openpyxl xlsxwriter pycurl
____________________________________________________________________________________________________________________

PLEASE RUN THE SCRIPT, by performing the following steps.

STEP 1. Open and  Run the script LATENCY SCRIPT.py on Spyder.

STEP 2. If no errors, continue from STEP 3. If errors occur, return to STEP 0. 

STEP 3. Input the number of iterations performed for each link within each archive list.
- For this case study, please perform 15 iterations, then another 15 iterations (total of 30 iterations).

(Script will run- The duration of the run will vary but could take anywhere from a few minutes for a small number of iterations, to hours for a large number of iterations.)
 
STEP 4. When the script run has completed, a '✓ Data saved to Excel: combined_latency_date_time.xlsx' message will appear. 

STEP 5. The output file combined_latency_date_time.xlsx will be present in script folder.
If not, it will aso be in file explorer of Spyder. Right click on the file name, click 'Show in folder', and send the output file to exk419@student.bham.ac.uk

*************************************************************************************************************
*PLEASE PERFORM THE SAME RUN ON A SIM CARD HAVING SUFFIECIENT AMOUNT OF BUNDLE DATA (20GB OF DATA OR MORE)	          *
*If less bundle data is available, please cut down iteration according refering to the data volume information below.		          *
*************************************************************************************************************
For any questions, please email me at exk419@student.bham.ac.uk
________________________________________________________________________________________________________________
Please note- Data Volumes for iterations are as follows:
1 Iteration- 2.16GB
10 iterations- 21.6 GB

Thank you for your participation.