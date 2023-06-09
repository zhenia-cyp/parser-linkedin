This parser collects the following information from LinkedIn:

a) Company name.
b) Company location.
c) Company website link.
e) link to the company's linkedin profile.

The parser opens company pages, extracts the specified data, and saves it in an Excel file. 
If the data is missing or the website link is inaccessible, the corresponding cells in the Excel file 
are marked in red.

The gathering of all the data occurs in the get_main_data() function using JavaScript and 
can also be easily expanded to extract more specific information.

To bypass certain detection and blocking mechanisms by LinkedIn, the selenium_stealth library is used, 
which has proven itself effective in practice.

1) Open the command prompt (terminal) on your computer.
Navigate to the folder where you want to save the repository and
use the command. 

     git clone https://github.com/zhenia-cyp/parser-linkedin.git

2) Use terminal navigate to the project folder where "ParserLinkedinPages.py", "requirements.txt"
and the "chromedriver" folder are located.

3) Check the version of your Google Chrome browser in the settings.

Go to the website https://chromedriver.chromium.org/downloads.

Download the appropriate version of ChromeDriver that matches your Google Chrome browser version.

Copy the chromedriver.exe file to the chromedriver folder.

In the enter_to_linkedin() function replace the file path with your own.

        driver = webdriver.Chrome(executable_path="C:\\Users\\josephbiden\\Desktop\\parser\\chromedriver\\chromedriver.exe")

Don't forget to use double '\\'.

4) To set your LinkedIn login and password in the enter_to_linkedin() function.
   
          driver.find_element(By.XPATH, '//*[@id="username"]').send_keys("Login to LinkedIn")
          driver.find_element(By.XPATH, '//*[@id="password"]').send_keys("password for LinkedIn")

5) In the enter_to_linkedin() function replace search link with filters to your own 
or keep the existing one for demonstration purposes

         driver.get("https://www.linkedin.com/sales/search/company?query=......")

6) By default, the parser will process only 2 pages. Please  specify the desired number of pages in the go_to_link() 
function.

        if not scroll_to_end_of_page.num_calls >= 2:
          scroll_to_end_of_page(driver)
        else:
            driver.quit()
            exit()

7) Replace "path = r"E:\result25.xlsx"" with the actual path and filename where you want to save the results.

8) Create a virtual environment use the terminal 

python -m venv venv

9) Activate a virtual environment 
 
 For Windows:
 venv\Scripts\activate

 For macOS and Linux:
 source myenv/bin/activate

10) Install dependencies

pip install -r requirements.txt 

11) Run the script

python ParserLinkedinPages.py
