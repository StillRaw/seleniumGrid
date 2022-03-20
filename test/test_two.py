import unittest
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import pandas as pd
import glob


class TestOne(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Remote(
            command_executor="http://192.168.1.29:4444/wd/hub",
            desired_capabilities={
                "browserName": "firefox",
            })
        self.driver.implicitly_wait(30)
        self.driver.maximize_window()

    def test_one(self):
        driver = self.driver
        global dict1
        dict1={"Name":[],
            "Year":[],
            "Origin":[],
            "Rating":[],
            "Production Company":[],
            "Budget":[],
            "Quotes":[],
            "Goofs":[]
            }
        
        for i in range (126,151):
            url = "https://www.imdb.com/chart/top/"
            driver.get(url)

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a")))
            movie_name  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a").text
            movie_year  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/span").text
            movie_rating  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[3]/strong").text

            driver.execute_script("arguments[0].click();", driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a"))
            
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/footer/div[2]/p")))

            movie_origin  = driver.find_element(by=By.XPATH, value="//*[contains(text(),' of origin')]/following-sibling::div[1]").text
            movie_production = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Production compan')]/following-sibling::div[1]").text
            movie_budget = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Budget')]/following-sibling::div[1]").text
            movie_quotes = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Quotes')]/following-sibling::div[1]").text
            movie_goofs = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Goofs')]/following-sibling::div[1]").text
            
            #adjusting texts
            movie_budget = movie_budget.strip(" (estimated)")#remove the estimated text
            movie_year = movie_year[1:5]#just get the year 1994 not (1994)
            movie_rating = movie_rating.replace(",",".")#Firefox using "," istead of "."

            
            dict1["Name"].append(movie_name)
            dict1["Year"].append(movie_year)
            dict1["Origin"].append(movie_origin)
            dict1["Rating"].append(movie_rating)
            dict1["Production Company"].append(movie_production)
            dict1["Budget"].append(movie_budget)
            dict1["Quotes"].append(movie_quotes)
            dict1["Goofs"].append(movie_goofs)

        df= pd.DataFrame.from_dict(dict1)
        file_name = '6.xlsx'
        sheet_name = 'IMDB Top 250'
        writer = pd.ExcelWriter("C:\\Users\\Onur\\Desktop\\iskur\\test\\"+file_name, engine='xlsxwriter')
        df.to_excel(writer, sheet_name=sheet_name, index = False)
        writer.save()
    
    def test_two(self):
        driver = self.driver
        global dict2
        dict2={"Name":[],
            "Year":[],
            "Origin":[],
            "Rating":[],
            "Production Company":[],
            "Budget":[],
            "Quotes":[],
            "Goofs":[]
            }
        
        
        for i in range (151,176):
            url = "https://www.imdb.com/chart/top/"
            driver.get(url)

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a")))
            movie_name  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a").text
            movie_year  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/span").text
            movie_rating  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[3]/strong").text

            driver.execute_script("arguments[0].click();", driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a"))
            
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/footer/div[2]/p")))

            movie_origin  = driver.find_element(by=By.XPATH, value="//*[contains(text(),' of origin')]/following-sibling::div[1]").text
            movie_production = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Production compan')]/following-sibling::div[1]").text
            movie_budget = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Budget')]/following-sibling::div[1]").text
            movie_quotes = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Quotes')]/following-sibling::div[1]").text
            movie_goofs = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Goofs')]/following-sibling::div[1]").text
            
            #adjusting texts
            movie_budget = movie_budget.strip(" (estimated)")#remove the estimated text
            movie_year = movie_year[1:5]#just get the year 1994 not (1994)
            movie_rating = movie_rating.replace(",",".")#Firefox using "," istead of "."

            
            dict2["Name"].append(movie_name)
            dict2["Year"].append(movie_year)
            dict2["Origin"].append(movie_origin)
            dict2["Rating"].append(movie_rating)
            dict2["Production Company"].append(movie_production)
            dict2["Budget"].append(movie_budget)
            dict2["Quotes"].append(movie_quotes)
            dict2["Goofs"].append(movie_goofs)

        df= pd.DataFrame.from_dict(dict2)
        file_name = '7.xlsx'
        sheet_name = 'IMDB Top 250'
        writer = pd.ExcelWriter("C:\\Users\\Onur\\Desktop\\iskur\\test\\"+file_name, engine='xlsxwriter')
        df.to_excel(writer, sheet_name=sheet_name, index = False)
        writer.save()
    
    def test_three(self):
        driver = self.driver
        global dict2
        dict2={"Name":[],
            "Year":[],
            "Origin":[],
            "Rating":[],
            "Production Company":[],
            "Budget":[],
            "Quotes":[],
            "Goofs":[]
            }
        
        
        for i in range (176,201):
            url = "https://www.imdb.com/chart/top/"
            driver.get(url)

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a")))
            movie_name  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a").text
            movie_year  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/span").text
            movie_rating  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[3]/strong").text

            driver.execute_script("arguments[0].click();", driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a"))
            
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/footer/div[2]/p")))

            movie_origin  = driver.find_element(by=By.XPATH, value="//*[contains(text(),' of origin')]/following-sibling::div[1]").text
            movie_production = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Production compan')]/following-sibling::div[1]").text
            movie_budget = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Budget')]/following-sibling::div[1]").text
            movie_quotes = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Quotes')]/following-sibling::div[1]").text
            movie_goofs = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Goofs')]/following-sibling::div[1]").text
            
            #adjusting texts
            movie_budget = movie_budget.strip(" (estimated)")#remove the estimated text
            movie_year = movie_year[1:5]#just get the year 1994 not (1994)
            movie_rating = movie_rating.replace(",",".")#Firefox using "," istead of "."

            
            dict2["Name"].append(movie_name)
            dict2["Year"].append(movie_year)
            dict2["Origin"].append(movie_origin)
            dict2["Rating"].append(movie_rating)
            dict2["Production Company"].append(movie_production)
            dict2["Budget"].append(movie_budget)
            dict2["Quotes"].append(movie_quotes)
            dict2["Goofs"].append(movie_goofs)

        df= pd.DataFrame.from_dict(dict2)
        file_name = '8.xlsx'
        sheet_name = 'IMDB Top 250'
        writer = pd.ExcelWriter("C:\\Users\\Onur\\Desktop\\iskur\\test\\"+file_name, engine='xlsxwriter')
        df.to_excel(writer, sheet_name=sheet_name, index = False)
        writer.save()
    
    def test_four(self):
        driver = self.driver
        global dict2
        dict2={"Name":[],
            "Year":[],
            "Origin":[],
            "Rating":[],
            "Production Company":[],
            "Budget":[],
            "Quotes":[],
            "Goofs":[]
            }
        
        
        for i in range (201,226):
            url = "https://www.imdb.com/chart/top/"
            driver.get(url)

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a")))
            movie_name  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a").text
            movie_year  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/span").text
            movie_rating  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[3]/strong").text

            driver.execute_script("arguments[0].click();", driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a"))
            
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/footer/div[2]/p")))

            movie_origin  = driver.find_element(by=By.XPATH, value="//*[contains(text(),' of origin')]/following-sibling::div[1]").text
            movie_production = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Production compan')]/following-sibling::div[1]").text
            movie_budget = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Budget')]/following-sibling::div[1]").text
            movie_quotes = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Quotes')]/following-sibling::div[1]").text
            movie_goofs = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Goofs')]/following-sibling::div[1]").text
            
            #adjusting texts
            movie_budget = movie_budget.strip(" (estimated)")#remove the estimated text
            movie_year = movie_year[1:5]#just get the year 1994 not (1994)
            movie_rating = movie_rating.replace(",",".")#Firefox using "," istead of "."

            
            dict2["Name"].append(movie_name)
            dict2["Year"].append(movie_year)
            dict2["Origin"].append(movie_origin)
            dict2["Rating"].append(movie_rating)
            dict2["Production Company"].append(movie_production)
            dict2["Budget"].append(movie_budget)
            dict2["Quotes"].append(movie_quotes)
            dict2["Goofs"].append(movie_goofs)

        df= pd.DataFrame.from_dict(dict2)
        file_name = '9.xlsx'
        sheet_name = 'IMDB Top 250'
        writer = pd.ExcelWriter("C:\\Users\\Onur\\Desktop\\iskur\\test\\"+file_name, engine='xlsxwriter')
        df.to_excel(writer, sheet_name=sheet_name, index = False)
        writer.save()

    def test_five(self):
        driver = self.driver
        global dict2
        dict2={"Name":[],
            "Year":[],
            "Origin":[],
            "Rating":[],
            "Production Company":[],
            "Budget":[],
            "Quotes":[],
            "Goofs":[]
            }
        
        for i in range (226,251):
            url = "https://www.imdb.com/chart/top/"
            driver.get(url)

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a")))
            movie_name  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a").text
            movie_year  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/span").text
            movie_rating  = driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[3]/strong").text

            driver.execute_script("arguments[0].click();", driver.find_element(by=By.XPATH, value="/html/body/div[3]/div/div[2]/div[3]/div/div[1]/div/span/div/div/div[3]/table/tbody/tr["+str(i)+"]/td[2]/a"))
            
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/footer/div[2]/p")))

            movie_origin  = driver.find_element(by=By.XPATH, value="//*[contains(text(),' of origin')]/following-sibling::div[1]").text
            movie_production = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Production compan')]/following-sibling::div[1]").text
            movie_budget = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Budget')]/following-sibling::div[1]").text
            movie_quotes = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Quotes')]/following-sibling::div[1]").text
            movie_goofs = driver.find_element(by=By.XPATH, value="//*[contains(text(),'Goofs')]/following-sibling::div[1]").text
            
            #adjusting texts
            movie_budget = movie_budget.strip(" (estimated)")#remove the estimated text
            movie_year = movie_year[1:5]#just get the year 1994 not (1994)
            movie_rating = movie_rating.replace(",",".")#Firefox using "," istead of "."

            
            dict2["Name"].append(movie_name)
            dict2["Year"].append(movie_year)
            dict2["Origin"].append(movie_origin)
            dict2["Rating"].append(movie_rating)
            dict2["Production Company"].append(movie_production)
            dict2["Budget"].append(movie_budget)
            dict2["Quotes"].append(movie_quotes)
            dict2["Goofs"].append(movie_goofs)

        df= pd.DataFrame.from_dict(dict2)
        file_name = '10.xlsx'
        sheet_name = 'IMDB Top 250'
        writer = pd.ExcelWriter("C:\\Users\\Onur\\Desktop\\iskur\\test\\"+file_name, engine='xlsxwriter')
        df.to_excel(writer, sheet_name=sheet_name, index = False)
        writer.save()
    
    
    def tearDown(self):
        self.driver.quit()
    

    


if __name__ == "__main__":
    unittest.main()

