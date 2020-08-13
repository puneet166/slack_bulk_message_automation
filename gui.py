from Tkinter import *
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pds
import os
import time
import pickle


window = Tk()
window.configure(background='red')


window.title("Slack bulk message automation")

window.geometry('720x720')

label1 = Label(window, text="Enter Slack Cookies ")

label1.grid(column=3, row=3,  padx=10, pady=10)

text1 = Entry(window,width=80)

text1.grid(column=8, row=3 , padx=10, pady=10)

label2 = Label(window, text="Enter Excel sheet path")

label2.grid(column=3, row=7,  padx=10, pady=10)

text2 = Entry(window,width=80)

text2.grid(column=8, row=7 , padx=10, pady=10)

label3 = Label(window, text="Message")

label3.grid(column=3, row=10 ,  padx=10, pady=10)

text3 = Entry(window,width=80)

text3.grid(column=8, row=10 ,  padx=10, pady=10)

def common(url,driver,message):
        newData = pds.read_excel(url) 
        #print(type(newData))
        name=newData.iloc[:,0].tolist()

        
            
        for i in range(0,len(name)):
            mes=message.replace("$name", name[i])
            time.sleep(1)

            test=driver.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[1]/div/div/button")
            time.sleep(1) 

            test.click()
            time.sleep(1)
            test1=driver.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[3]/div/div[2]/div/div/div[1]/div/div/div[1]/div/input")
            time.sleep(1) 
        
            test1.send_keys('@',name[i])

            time.sleep(1) 

            test1.send_keys(Keys.ENTER)

            time.sleep(1) 

            test2 = driver.find_element_by_tag_name('p')

            time.sleep(1) 

            test2.send_keys(mes)
            time.sleep(1) 
            test2.send_keys(Keys.ENTER)
        driver.quit


def clicked():

    cookies=text1.get()
    message = text3.get()
    url=text2.get()
    if len(cookies) != 0 and len(message) != 0:  # empty!
        if os.path.exists(url)==True:

            url=text2.get()
            driver =webdriver.Chrome()
            driver.get('https://app.slack.com/client/T0180SCUS23/C018MQ9M40L')
            for cookie in cookies:
                driver.add_cookie(cookie)
            driver.get('https://app.slack.com/client/T0180SCUS23/C018MQ9M40L')
            time.sleep(5)
            #files =('C:\Users\Puneet Singh\Documents\data1.xlsx') 
            
            common(url,driver,message)

            
        else:
            label5 = Label(window, text="path of excel file is incorrect")

            label5.grid(column=3, row=17,  padx=10, pady=10)
    else:
        label8 = Label(window, text="first fill the boxes")

        label8.grid(column=3, row=15,  padx=10, pady=10)

        

    
def create_session():
    
    message = text3.get()
    url=text2.get()
    if len(message) != 0 and len(url) != 0:
        if os.path.exists(url)==True:
            
            label6 = Label(window, text="After login, enter any no here for further processing on console ")


            label6.grid(column=3, row=25,  padx=10, pady=10)
            

            
            driver =webdriver.Chrome()

            driver.get("https://app.slack.com/client/T0180SCUS23/C018MQ9M40L")
            raw_input("press any key")
            time.sleep(5)
            pickle.dump( driver.get_cookies() , open("cookies.pkl","wb"))
            time.sleep(5)
            driver.quit()
            time.sleep(5)

            driver = webdriver.Chrome() # or add to your PATH

            driver.get('https://app.slack.com/client/T0180SCUS23/C018MQ9M40L')
            time.sleep(5)

            cookies = pickle.load(open("cookies.pkl", "rb"))
            time.sleep(5)

            for cookie in cookies:
                driver.add_cookie(cookie)
            time.sleep(5)
            driver.get('https://app.slack.com/client/T0180SCUS23/C018MQ9M40L')
            common(url,driver,message)

            
            
        else:
            label5 = Label(window, text="path of excel file is incorrect")

            label5.grid(column=3, row=20,  padx=10, pady=10)
    else:
        label8 = Label(window, text="first fill the boxes instead of cookies")

        label8.grid(column=3, row=22,  padx=10, pady=10)



    
btn = Button(window, text="submit", command=clicked)

btn.grid(column=15, row=18)

btn1 = Button(window, text=" press for Create slack session", command=create_session)

btn1.grid(column=15, row=10)


window.mainloop()