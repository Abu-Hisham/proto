# **PROTO EMAIL AUTOMATION** 

## How the program works 

This simple python script performs the following tasks

-  Fetches data from the database

- Opens an Excel Workbook

- Creates a new Spreadsheet (Names the spreadsheet using the Time stamp DD-MM-YYYY H-M-S)

- Appends the Data to the new sheet 

- Saves the workbook and emails it automatically to the specified recipients (in the program)

  

## Program Configurations 

For the program to run successfully and produce the desired output the below configurations need to be effected

1. Open the Script file automated_email and follow the instructions included in the comments
2. Edit the suggested areas as per your context (Machine Setting) in order to obtain the desired output

## Program Installation/Running Requirements

1. [Python]: https://www.python.org/downloads/

   version 2.7.x (or higher)* 

2. *MS Office (Particularly MS Excel)*

3. *Windows version 8.1 (or higher)*

## Running Instructions

1. Open Windows CMD (Hit **Windows** + R)

2. Clone this repository locally (preferably on your desktop) by typing then hitting **Enter** key

   ```
   git clone https://github.com/Abu-Hisham/proto.git 
   ```

3. Navigate into the repository by typing then hitting **Enter** key

   ```
   cd proto
   ```

4. install  virtual Environment(Used for installing dependencies locally in project folders) typing and hitting **Enter** Key

   ```
   pip install virtualenv 
   ```

5. Create your virtual environment by typing the below command and hitting **Enter** Key

   ```
   virtualenv venv
   ```

6. Activate the virtual environment by typing the below command and hitting **Enter**  Key

   ```
   venv\Scripts\activate
   ```

7. Install the program dependencies by typing the below command and hitting **Enter** Key

   ```
   pip install -r requirements.txt --no-index
   ```

8. Run the program by typing the following command and hitting the **Enter** Key

   ```
   python src/automated_email.py
   
   ```



 