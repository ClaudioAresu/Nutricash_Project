
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from dateutil.relativedelta import relativedelta
import datetime
import time
import pandas
import os

def cal_limit(limite, saldo):
    print ("SALDO ANTIGO: " + str(saldo))
    print ("LIMITE ANTIGO: " + str(limite))

    hoje = datetime.datetime.now()
    ontem = (hoje - datetime.timedelta(days=1)).day
    total_dias_mes = (((hoje + relativedelta(months=1)).replace(day=1)) - datetime.timedelta(days=1)).day

    permitido_cal = round(((limite - saldo) / ontem) * total_dias_mes, 2)

    if permitido_cal % 50 == 0:
        permitido_arr = permitido_cal
    else:
        permitido_arr = permitido_cal + (50 - (permitido_cal % 50))

    if permitido_arr > limite:
        dif_limite_cota = permitido_arr - limite
    else:
        print ("cota nao permitida")
        return 0

    novo_saldo = saldo + dif_limite_cota

    if permitido_arr > limite * 2:
        limite_final = limite * 2
    else:
        limite_final = permitido_arr

    print ("NOVO SALDO: " + str(novo_saldo))
    print ("LIMITE FINAL: " + str(limite_final))

    return '{:,.2f}'.format(limite_final)

def configs(headless):
    chromeOps = webdriver.ChromeOptions()
    #chromeOps._binary_location = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\Chrome.exe"
    chromeOps._binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\Chrome.exe"
    chromeOps._arguments = ["--enable-internal-flash"]
    if str(headless).lower() != 'sim':
        chromeOps.add_argument('--headless')

    return webdriver.Chrome("ChromeDriver_32\\chromedriver.exe", port=4445, options=chromeOps)

def main_url():
    return "https://gis.maxifrota.com.br"

def login(browser, user, password):
    browser.get(main_url())
    browser.find_element_by_name('j_username').send_keys(user)
    browser.find_element_by_name('j_password').send_keys(password)
    browser.find_element_by_xpath('//*[@id="loginForm"]/div[3]/button').click()

def change_card_limit(browser, new_limit, text):
    time.sleep(2)
    browser.find_element_by_xpath('//*[@id="page-content"]/div[1]/div[3]/form/div[1]/div/div[3]/div[4]/div/a').click()
    browser.find_element_by_id('novoLimite').clear()
    browser.find_element_by_id('novoLimite').send_keys(new_limit)
    browser.find_element_by_id('novolimiteInicial').clear()
    browser.find_element_by_id('novolimiteInicial').send_keys(new_limit)
    browser.find_element_by_name('motivo').send_keys(text)
    #browser.find_element_by_xpath('').click()


    return 0

def limits(browser, PLACA):
    time.sleep(5)
    #browser.find_element_by_xpath('/html/body/div[1]/div/header/div/div[2]/a/i').click()
    browser.find_element_by_xpath('//*[@id="ul-favorito-lateral"]/li[1]').click()
    time.sleep(5)
    browser.find_element_by_name('cartao').clear()
    browser.find_element_by_name('cartao').send_keys(PLACA)
    browser.find_element_by_id('idProduto_chosen').click()
    browser.find_element_by_xpath('/html/body/div[1]/div/div[1]/div[2]/div/div/div/div[3]/div/form/div/div/div/div[1]/div/div[2]/div/div/div/div/div/input').send_keys("ABASTECIMENTO")
    browser.find_element_by_xpath('/html/body/div[1]/div/div[1]/div[2]/div/div/div/div[3]/div/form/div/div/div/div[1]/div/div[2]/div/div/div/div/div/input').send_keys(Keys.ENTER)
    browser.find_element_by_xpath('//*[@id="page-content"]/div/div[3]/div/form/div/div/div/div[2]/div/div/div/a[1]').click()
    time.sleep(5)

    try:
        table = browser.find_element_by_xpath('/html/body/div[1]/div/div[1]/div[2]/div/div/div/div[4]/div/div[3]/div[5]/div[2]/table/tbody')
        table = table.find_elements_by_tag_name('tr')
    except TypeError as error:
        raise error

    for row in table:
        if "Cancelado" in row.text:
            pass
        else:
            items = row.find_elements_by_tag_name('td')
            values = [x.text for x in items]
            row.click()
            try:
                return float(values[8].replace(".","").replace(",",".")), float(values[9].replace(".","").replace(",","."))
            except Exception:
                raise TypeError
    

def main(browser, placas, text):
    for placa in placas:
        try:
            limit, sale = limits(browser, placa)
            new_limit = cal_limit(sale, limit)
            if new_limit != 0:
                change_card_limit(browser,new_limit,text)
                log_file.writelines(str(limit) + '\t' + str(new_limit) + '\t' + 'OK' +'\t' + placa + '\n')
                log_file.flush()
            else:
                log_file.writelines(str(limit) + '\t' + '0' + '\t' + 'NAO DISPONIVEL'+ '\t' + placa +'\n')
                log_file.flush()
        except TypeError as a:
            log_file.writelines('0' + '\t' + '0' + '\t' + 'NAO DISPONIVEL' + '\t' + placa + '\n')
            log_file.flush()
            continue
        except Exception as e:
            raise e

def log(path_df):
    global log_file
    base_dir = os.path.join('output',path_df.split('\\')[-2])
    if not os.path.exists(base_dir):
        try:
            os.makedirs(base_dir)
        except Exception as exc: 
            raise exc

    log_file = open(os.path.join(base_dir, path_df.split('\\')[-1].replace('xlsx', 'csv')), 'w')
    log_file.writelines('LIMITE\tNOVO_LIMITE\tSTATUS\tPLACA'+'\n')
    log_file.flush()

    return log_file

def read_confis():
    file = open("configure.txt", "r").readlines()

    user = file[0].split("=")[1].strip()
    password = file[1].split("=")[1].strip()
    headless = file[2].split("=")[1].strip()
    path_df = []

    for root, dirs, files in os.walk('input/'):
        for name in files:
            if name.endswith((".xlsx", ".xlsx")):
                path_df.append(os.path.join(root.replace('/','\\'),name))

    return user, password, headless, path_df

def load_df(path_df):
    global log_file

    df = pandas.read_excel(path_df, header=None)
    text = df[1].tolist()[5]
    placas = [x for x in df[6].tolist() if not type(x) is float]

    log_file = log(path_df)

    return text, placas


##########################################

user, password, headless, path_df = read_confis()
browser = configs(headless)
login(browser, user, password)
for df in path_df:
    text, placas = load_df(df)
    main(browser, placas, text)
print("END")
browser.close()