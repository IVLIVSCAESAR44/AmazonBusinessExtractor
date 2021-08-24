from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from datetime import timedelta
from datetime import date
from datetime import datetime
import pyodbc
import shutil

 #password and username to be used to log in
usernamestr = 'Username'
passwordstr = 'Password'
options = webdriver.ChromeOptions() 
options.add_argument("user-data-dir=Location of Chrome User data directory")
options.add_experimental_option("prefs", {"download.default_directory": r"Location where you want the file to go",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
  })
options.add_experimental_option("excludeSwitches", ['enable-automation'])

w = webdriver.Chrome(options=options)


try:
     w.get('https://vendorcentral.amazon.com')

   #do the actual filling in
     username = w.find_element_by_id('ap_email')
     username.click()
     username.clear()
     username.send_keys(usernamestr)
     time.sleep(1)

     password = w.find_element_by_id('ap_password')
     password.send_keys(passwordstr)
     time.sleep(1)

     signinbtn = w.find_element_by_id('signInSubmit')
     signinbtn.click()
     time.sleep(2)
     print('Sign in Complete....')
    
     reportshover = w.find_element_by_id("vss_navbar_tab_reports")
     hover = ActionChains(w).move_to_element(reportshover)
     hover.perform()

     analyticsbtn = w.find_element_by_xpath("//*[@id='ARAP_I90_Amazon_Retail_Analytics_text']").click()
     print('Reports tab clicked....')
     time.sleep(2)
    
     salesdiagbtn = w.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div/ul/li[1]/div/div/div/span[1]/a')
     salesdiagbtn.click()
     time.sleep(2)
   
     dropdownbtn = w.find_element_by_id('dashboard-filter-programToggler')
     dropdownbtn.click()
    
     amzbizbtn = w.find_element_by_xpath('//*[@id="dashboard-filter-programToggler"]/div/awsui-button-dropdown/div/div/ul/li[4]')
     amzbizbtn.click()
     time.sleep(2)
    
     salesviewbtn = w.find_element_by_id('dashboard-filter-viewFilter')
     salesviewbtn.click()
    
     shippedrevbtn = w.find_element_by_xpath('//*[@id="dashboard-filter-viewFilter"]/div/awsui-button-dropdown/div/div/ul/li[2]')
     shippedrevbtn.click()
    
     applybtn = w.find_element_by_xpath('//*[@id="dashboard-filter-applyCancel"]/div/awsui-button[2]/button')
     applybtn.click()
     time.sleep(4)
    
     addbtn = w.find_element_by_id('dashboard-filter-visCols')
     addbtn.click()
    
     upcbtn = w.find_element_by_xpath('//*[@id="dashboard-filter-visCols"]/div/awsui-button-dropdown/div/div/ul/li[4]')
     upcbtn.click()
    
     addbtn.click()
     eanbtn = w.find_element_by_xpath('//*[@id="dashboard-filter-visCols"]/div/awsui-button-dropdown/div/div/ul/li[3]')
     eanbtn.click()
    
     addbtn.click()
     brandbtn = w.find_element_by_xpath('//*[@id="dashboard-filter-visCols"]/div/awsui-button-dropdown/div/div/ul/li[5]')
     brandbtn.click()
    
     addbtn.click()
     bindingbtn = w.find_element_by_xpath('//*[@id="dashboard-filter-visCols"]/div/awsui-button-dropdown/div/div/ul/li[11]')
     bindingbtn.click()
    
     addbtn.click()
     modelbtn = w.find_element_by_xpath('//*[@id="dashboard-filter-visCols"]/div/awsui-button-dropdown/div/div/ul/li[13]')
     modelbtn.click()

     downloadbtn = w.find_element_by_id('downloadButton')
     downloadbtn.click()
     time.sleep(5)
    
     csvbtn = w.find_element_by_xpath('//*[@id="downloadButton"]/awsui-button-dropdown/div/div/ul/li[3]/ul/li[2]')
     csvbtn.click()
    
     WebDriverWait(w, 60).until(EC.alert_is_present())         
     alert = w.switch_to.alert
     alert.accept()
     print('Alert accepted....')
   
     time.sleep(10)
     logoutbtn = w.find_element_by_id('logout_topRightNav')
     logoutbtn.click()
     w.quit()
     print('Webdriver portion complete....')
    
     today = date.today()
     subDays = timedelta(4)
    
     execdate = today - subDays
     strexecdate = str(execdate)   
    
#Delete first row
     path = r'Location of stored filed from download'
     print('File Opened....')
    
     new_df = pd.read_csv(path + 'Sales Diagnostic_Detail View_US.csv', sep=',', error_bad_lines=False, dtype ='str', encoding='utf-8-sig',header=1)

     new_df.drop(new_df.filter(regex='Unnamed'),axis=1, inplace=True)
     new_df.columns = new_df.columns.str.replace(' ','')    
     new_df.columns = new_df.columns.str.replace('-','')
     new_df.columns = new_df.columns.str.replace('/','')
     new_df.columns = new_df.columns.str.replace('%','')
     new_df.columns = new_df.columns.str.replace('(','')
     new_df.columns = new_df.columns.str.replace(')','')



     print('Formatting column names for data extract....')
    
     new_df = new_df.replace(',','', regex=True)
     new_df = new_df.replace('—','', regex=True)
     new_df = new_df.fillna(value=0)
        
     new_df["ActivityDate"] = execdate
     print('Adding Activity Date Column')
    
     conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};'
                          'Server=FWS-InsertServer;'
                          'Database=InsertDatabase;'
                          'Trusted_Connection=yes;'
                          'autocommit=False;')
    
     now = datetime.now()
     current_time = now.strftime("%H:%M:%S")
    
     print('Beginning Data Load: ' + current_time)
     cursor=conn.cursor()
     for row in new_df.itertuples():
         cursor.execute('''
                     INSERT INTO dbo.[Table] 
             ([ASIN]
              ,[EAN]
              ,[UPC]
              ,[ProductTitle]
              ,[Brand]
              ,[Binding]
              ,[Model StyleNumber]
              ,[ShippedRevenue]
              ,[ShippedRevenue%ofTotal]
              ,[ShippedRevenuePriorPeriod]
              ,[ShippedRevenueLastYear]
              ,[ShippedUnits]
              ,[ShippedUnits%ofTotal]
              ,[ShippedUnitsPriorPeriod]
              ,[ShippedUnitsLastYear]
              ,[Subcategory(SalesRank)]
              ,[AverageSalesPrice]
              ,[AverageSalesPricePriorPeriod]
              ,[ChangeinGlanceViewPriorPeriod]
              ,[ChangeinGVLastYear]
              ,[RepOOS]
              ,[RepOOS%ofTotal]
              ,[RepOOSPriorPeriod]
              ,[LBB(Price)]
              ,[ActivityDate])
            
             VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
             ''',
             row.ASIN,
             row.EAN,
             row.UPC,
             row.ProductTitle,
             row.Brand,
             row.Binding,
             row.ModelStyleNumber,
             row.ShippedRevenue,
             row.ShippedRevenueofTotal,
             row.ShippedRevenuePriorPeriod,
             row.ShippedRevenueLastYear,
             row.ShippedUnits,
             row.ShippedUnitsofTotal,
             row.ShippedUnitsPriorPeriod,
             row.ShippedUnitsLastYear,
             row.SubcategorySalesRank,
             row.AverageSalesPrice,
             row.AverageSalesPricePriorPeriod,
             row.ChangeinGlanceViewPriorPeriod,
             row.ChangeinGVLastYear,
             row.RepOOS,
             row.RepOOSofTotal,
             row.RepOOSPriorPeriod,
             row.LBBPrice,
             row.ActivityDate
             )
     now1 = datetime.now()
     current_time1 = now1.strftime("%H:%M:%S")
     print('load complete: ' + current_time1)                
     cursor.commit()
     cursor.close()
     conn.close()    
     print('Connection closed')    
    
     new_df.to_csv(path + 'Sales Diagnostic_Detail View_US.csv', index=False, encoding='utf-8-sig')
     print('File Saved')
    
     shutil.move('Old File Location with file name', 'Where you want to archive the file')
     print('File moved to Archive folder')
    
except Exception as e:
    print(e)

