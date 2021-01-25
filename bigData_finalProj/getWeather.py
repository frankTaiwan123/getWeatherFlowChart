import os
import time
from selenium import webdriver
from selenium.webdriver.support.ui import Select

yearlist = [2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019]
#yearlist = [2016, 2017, 2018, 2019]

for Years in yearlist:
    print(Years)
    
    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': 'D:\\Main\\NTUST\\bigData_finalTeamProj\\weather\\'+str(Years)}
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(
        executable_path="D:\\Main\\NTUST\\bigData_finalTeamProj\\chromedriver79.exe",
        chrome_options=options)
    
    driver.get("https://e-service.cwb.gov.tw/HistoryDataQuery/index.jsp")
    assert "觀測資料查詢系統" in driver.title
    
    driver.implicitly_wait(30)
    selectCountry = Select(driver.find_element_by_id('stationCounty'))
    for city in range(len(selectCountry.options)):
        currCity = selectCountry.options[city].text
        print(currCity)
        selectCountry.select_by_index(city)
        driver.implicitly_wait(30)
        
        selectStation = Select(driver.find_element_by_id('station'))
        for station in range(len(selectStation.options)):
            currStation = selectStation.options[station].text
            print("\t"+currStation)
            if "撤銷站" in currStation:
                print("廢棄資料")
                continue
            
            fileExist = False
            dir_link = 'D:\\Main\\NTUST\\bigData_finalTeamProj\\weather\\'+str(Years)
            newName = currCity.split(" ")[0]+"_"+currStation.split(" ")[0]+"_"+str(Years)+".csv"
            if "?" in newName:
                newName = newName.replace("?", "[E]")
            for file in os.listdir(dir_link+"\\"+currCity+"\\"):
                if newName in file:
                    print(currCity.split(" ")[0]+"_"+currStation.split(" ")[0]+"_"+str(Years)+".csv"+" exist!")
                    fileExist = True
                    break

            if fileExist:
                continue
            
            selectStation.select_by_index(station)
            driver.implicitly_wait(30)
    
            selectClass = Select(driver.find_element_by_id('dataclass'))
            selectClass.select_by_value("inquire")
            driver.implicitly_wait(30)
    
            selectType = Select(driver.find_element_by_id('datatype'))
            selectType.select_by_value("year")
            driver.implicitly_wait(30)
    
            year = driver.find_element_by_id("datepicker")
            year.clear()
            year.send_keys(str(Years))
    
            driver.find_element_by_id("doquery").click()
            driver.implicitly_wait(30)
            time.sleep(1)
    
            windows=driver.window_handles
            driver.switch_to.window(windows[-1])
            
            driver.find_element_by_css_selector('a#downloadCSV>input').click()
            time.sleep(2)
            driver.close()
            driver.switch_to_window(windows[0])
            
            dir_link = 'D:\\Main\\NTUST\\bigData_finalTeamProj\\weather\\'+str(Years)
            dir_lists = os.listdir(dir_link)
            dir_lists.sort(key=lambda fn: os.path.getmtime(dir_link + '\\' + fn))
            os.chdir(dir_link)
            newName = currCity.split(" ")[0]+"_"+currStation.split(" ")[0]+"_"+str(Years)+".csv"
            if "?" in newName:
                newName = newName.replace("?", "[E]")
            os.rename(dir_lists[-1], newName)
            os.replace(dir_link+"\\"+newName, dir_link+"\\"+currCity+"\\"+newName)
            
        print()
        
    driver.quit()
    print()
