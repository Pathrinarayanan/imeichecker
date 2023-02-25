from flask import Flask, render_template, request, send_file
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
import datetime

now = datetime.date.today()
app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html', message='welcome imei checker!')

@app.route('/upload', methods=['POST'])
def upload():
    try:
        listval=[]
        file = request.files['file']
        #df = pd.read_excel(file)
        #df = df.iloc[:, 1:] # Remove the first column
        wb = load_workbook(file)
        sheet = wb.active
    except:
        return render_template('index.html',  message='Please upload a file in .xlsx format!')
    val =""

    driver = webdriver.Chrome('E:\\chromedriver')
    driver.get("https://www.cyberyodha.org/2022/10/imei-bulk-lookup-tool.html")
    search_bar = driver.find_element(By.XPATH, '//*[@id="imei"]/textarea')
    sheet1 = sheet['A']

    for data in sheet1:
        print(data.value)
        search_bar.clear()
        try:
            search_bar.send_keys(int(data.value))
            #search_bar.send_keys(Keys.RETURN)
            button = driver.find_element(By.XPATH,'//*[@id="submit"]')
            button.click()
            driver.implicitly_wait(15)
            
        except:
            listval.append("malformed value")
            continue
        lv = driver.find_elements(By.XPATH, '//*[@id="imeiTable"]/tr[2]/td[3]')
        if lv == []:
            listval.append('unknown')
        try:
            for element in driver.find_elements(By.XPATH, '//*[@id="imeiTable"]/tr[2]/td[3]'):
                if element.is_displayed():
                    val = element.text
                    print(val)
                
                    
                    listval.append(val)
                else:
                    listval.append("unknown")
        except:
            listval.append("unknown")
        
        #listval.append("malformed value")
    listval=listval[1:]
    print(listval)

    df = pd.read_excel(file)
    df['model'] = listval


    modified_file =  str(now) +"_" + file.filename
    df.to_excel(modified_file, index=False)
    driver.close()
    return render_template('download.html', value=modified_file)
    

@app.route('/download')
def download():
    filename = request.args.get('filename')
    return send_file(filename, as_attachment=True)

if __name__ == '_main_':
    app.run(debug=True)
