from tkinter import *
from docxtpl import *
from tkinter import messagebox
from docxtpl import DocxTemplate

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common import utils
from selenium.webdriver import support
from pywinauto.application import Application
import time
import datetime



def Open_Word(ev):
    os.system("start winword.exe SEZ_Template_out.docx")
    return

def Quit(ev):
    global root
    root.destroy()

def Write1(ev):
    print('start')

    print(' dictionary {}'.format(dictionary))
    print('Словарь vr = {}'.format(vr))

    for r1, elem1 in enumerate(dictionary):
        print('r1='+str(r1))
        print ('elem1='+str(elem1))
        
      
        if elem1=='config':

                tx=vr[elem1].get("1.0",END)
                print('tx = {}'.format(tx))
                dictionary[elem1][1]=R(tx)

        else:
               dictionary[elem1][1]=vr[elem1].get()
       
    print ('Записали  {}'.format(dictionary))
    write_to_Word()
    return()

    #Функция проверки наличия диалогового окна
def test_alert():
    try:
        alert = driver.switch_to.alert
        print('Есть диалоговое окно')
        return True
    except :
        print('Диалогового окна нет')
        return False
    #**************************************************

def send_to_Sez(ev):

    os.chdir(sys.path[0])
    from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
    """
    from winreg import *

    def Enable_Protected_Mode():
        # SECURITY ZONES ARE AS FOLLOWS:
        # 0 is the Local Machine zone
        # 1 is the Intranet zone
        # 2 is the Trusted Sites zone
        # 3 is the Internet zone
        # 4 is the Restricted Sites zone
        # CHANGING THE SUBKEY VALUE "2500" TO DWORD 0 ENABLES PROTECTED MODE FOR THAT ZONE.
        # IN THE CODE BELOW THAT VALUE IS WITHIN THE "SetValueEx" FUNCTION AT THE END AFTER "REG_DWORD".
        #os.system("taskkill /F /IM iexplore.exe")
        try:
            keyVal = r'Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1'
            key = OpenKey(HKEY_CURRENT_USER, keyVal, 0, KEY_ALL_ACCESS)
            SetValueEx(key, "2500", 0, REG_DWORD, 0)
            print("enabled protected mode")
        except Exception:
            print("failed to enable protected mode")
    """



    caps = DesiredCapabilities.INTERNETEXPLORER
    caps['ignoreProtectedModeSettings'] = True


    print ('step 1')
    #Enable_Protected_Mode()


    #driver = webdriver.Ie()
    driver = webdriver.Ie(capabilities=caps)

    #driver = webdriver.Chrome()
    print ('step 2')
    url_sez="http://helpdesk/manager/-/?ac%3Dshowtask%26flag%3D%26pid%3D229%26id%3D642%26task%3D"+dictionary['num_sez'][1]
    print('url = {}'.format(url_sez))
#    driver.get ("http://helpdesk/manager/-/?ac%3Dshowtask%26flag%3D%26pid%3D229%26id%3D642%26task%3D478106")
    driver.get(url_sez)
    print ('step 3')

    '''
    elem = driver.find_element_by_link_text("Поиск").click()
    elem2=driver.find_element_by_name("task")
    num_task='478106'
    elem2.send_keys(num_task)
    elem2.send_keys(Keys.RETURN)

    time.sleep(3)
    '''
    
    driver.find_element_by_name("theme").send_keys("Прошу ")


    #Проеряем наличие всплывающего окна
 #   time.sleep(2)

    time.sleep(3)
    try:
        driver.switch_to.alert
        print('Есть диалоговое окно')
        Alert(driver).accept()

    except:
        print('Диалогового окна нет')

    """        
    if test_alert():
    #    Alert=driver.switch_to_alert
        print('щелкаем ок')
        Alert(driver).accept()
    """
    driver.find_element_by_name("theme").send_keys("согласовать изменение")
    txt_details=dictionary["bizcel"][0]+' : '+dictionary["bizcel"][1]+'\n'+\
        dictionary["systemname"][0]+' : '+dictionary["systemname"][1]+'\n'+\
        dictionary["vliyanie_klients"][0]+ ' : '+dictionary["vliyanie_klients"][1]+'\n'+\
        dictionary["srok"][0]+ ' : ' + dictionary["srok"][1] +'\n'+\
        dictionary["dlitelnost"][0]+ ' :' +dictionary["dlitelnost"][1]
    

    
    driver.find_element_by_name("text").send_keys(txt_details)
    
    driver.find_element_by_link_text("Вложить файл").click()
    #driver.find_element_by_link_text("Обзор...").click()
    driver.find_element_by_name("attach[]").send_keys(sys.path[0].replace('/','\\')+'\\SEZ_Template_out.docx')
#    driver.find_element_by_name("attach[]").send_keys("C:\\SEZ_Template_out.docx')
    
    alert=driver.switch_to_alert

    print("******")

    l=driver.window_handles
    print(" L[0]= {}".format(l[0]))
    app = Application().connect(path=r"C:\Program Files\Internet Explorer\iexplore.exe")

    """
    dlg = app.top_window()
    print(dlg)
    dialogs = app.windows()
    print("dialogs  {}".format(app.windows()))

    print ("title {}".format(app.window(title_re=".*Список пользователей.*")))

    print("!!!!!")
    print(l)
    """
    print( 'alert  {0}'.format(Alert.text))



    print ('step 3 Кликаем на "Назначить ответственного"')
    driver.find_element_by_link_text("Назначить ответственного").click()

    time.sleep(3)
    #Переключаемся на всплывающее окно


    l=driver.window_handles

    print(l)
    print(len(l))
    if len(l)>1:
        driver.switch_to.window(l[1])
        print("переключились на всплывающее окно")

    #if test_alert():
    try :

          driver.find_element_by_xpath("//input[@name='search']")
          print('Поиск успешный  {}'.format(dictionary["soglasovatel"][1]))

              
    except:
          print("Поиск не успешный, Переключаемся на другое окно")
          driver.switch_to.window(l[0])
          
    time.sleep(3)
    #input("Press Enter to continue...")
    find_name=dictionary["soglasovatel"][1]
    inp=driver.find_element_by_xpath("//input[@name='search']")
    inp.click()    
    print('Отправляем  ФИО {}'.format(dictionary["soglasovatel"][1]))

    inp.send_keys(find_name)
    webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()
    time.sleep(2)
#    for_find=dictionary["soglasovatel"][1]
    driver.find_element_by_link_text(find_name).click()
     # print("end")
#    driver.switch_to.window(l[1])
#    time.sleep(3)
    driver.switch_to.window(l[0])

    try:
        print('Пытаемся сохранить')
        driver.find_element_by_xpath("//button['Сохранить изменения']")
    except:
        print('Не получилось переключаем окно и сохраняем')
        driver.switch_to.window(l[1])

        
    driver.find_element_by_xpath("//button['Сохранить изменения']").click()

    print ('end')
              


#********end send_to_Sez*************

def write_to_Word():

    print('start to write in Word')
    doc = DocxTemplate("SEZ_Template.docx")
    print('open ok')
    context={}


    """
    num_sez="45985"
    zakazchik="Кусков А.В."
    kurator = "Жильцов А.П."
    bizcel = "Доступ с сервера мониторинга оборудования IBM на оборудование СХД"
    ispolnitel ="Мещеряков С.С."
    systemname = "Файрвол Cisco ASA cdm-fw-core  через который осуществляется доступ."
    testirovanie ="Неприменимо"
    config = '''access-list mgmt_in line 2 remark *** ipi 471991 ***

    access-list mgmt_in line 3 extended permit ip any host 172.16.25.29 (hitcnt=0) 0x33191bfd
    access-list mgmt_in line 4 extended permit ip any host 172.16.25.27 (hitcnt=0) 0x71874c12
    access-list mgmt_in line 5 extended permit ip any host 172.16.25.28 (hitcnt=0) 0x7a0459a0
    access-list mgmt_in line 6 extended permit ip any host 10.99.209.60 (hitcnt=1) 0xc2c73cf0
    access-list mgmt_in line 7 extended permit ip any host 10.99.209.64 (hitcnt=0) 0x439293cb
    access-list mgmt_in line 8 extended permit ip any host 10.99.209.65 (hitcnt=0) 0xb0921744
    access-list mgmt_in line 9 extended permit ip any host 10.99.209.52 (hitcnt=0) 0x6f30850f
    access-list mgmt_in line 10 extended permit ip any host 10.99.209.53 (hitcnt=0) 0xf75a3e82
    access-list mgmt_in line 11 extended permit ip any host 10.99.209.51 (hitcnt=0) 0xa681bb9a
    access-list mgmt_in line 12 extended permit ip any host 10.99.209.50 (hitcnt=0) 0x597d0f4f
    access-list mgmt_in line 13 extended permit ip any host 10.77.200.52 (hitcnt=0) 0xc290ad58
    '''
    vliyanie_klients='Изменения влияния не оказывают'
    vremya_otkaza_klients='Отсутствует.'
    obhodnie_puti_klients = 'Неприменимо.'
    vliyanie_polzovatel ='Изменения влияния не оказывают.'
    vremya_otkaza_polzovatel = 'Нет'
    obhodnie_puti_polzovatel = 'Нет'
    otkat='Удаление из конфигурации строк:'
    srok ='10.07.2018  11:00'
    dlitelnost = '10 мин'
    
    

    context = { 'num_sez' : num_sez,
                'zakazchik' : zakazchik,
                'kurator': kurator,
                'bizcel': bizcel,
                'ispolnitel': ispolnitel,
                'systemname': systemname,
                'testirovanie' : testirovanie,
                'config' : config,
                'vliyanie_klients' : vliyanie_klients,
                'vremya_otkaza_klients' : vremya_otkaza_klients,
                'obhodnie_puti_klients' : obhodnie_puti_klients,
                'vliyanie_polzovatel' : vliyanie_polzovatel,
                'vremya_otkaza_polzovatel' : vremya_otkaza_polzovatel,
                'obhodnie_puti_polzovatel' : obhodnie_puti_polzovatel,
                'otkat' : otkat,
                'srok' : srok,
                'dlitelnost' : dlitelnost
                }

    """

    for r2, elem2 in enumerate(dictionary):

              context[elem2]=dictionary[elem2][1]  

    print('context =  {}'.format(context))
    now = datetime.datetime.now()
    context['data_sozd']=now.strftime("%d/%m/%Y")
    doc.render(context)

    print ('insert ok')
    #context2 ={ 'var_name2' : "Привет Тест2!" }
    #doc.render(context2)
    print('render ok')
    doc.save("SEZ_Template_out.docx")
    print ('save ok')

    messagebox.showinfo("Успешная запись", "Создан файл изменения SEZ_Template_out.docx")

    
    return()
    

#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
if __name__ == '__main__':

#Открываем конфиг файл и создаем словарь

    os.chdir(sys.path[0].replace('/','\\'))
    f = open("config_change_process.cfg")
    global dictionary
    dictionary={}
    for line in f.readlines():
    #    print (line)
        spisok=line.strip().split("=")
        print('spisok = {}'.format(spisok))
        for w, u in enumerate (spisok):
              spisok[w]=u.strip()        
        dictionary[spisok[0]]=spisok[1:]

    

    print ('dictionary = {}'.format(dictionary))
    f.close()


    mes="MES"
    root = Tk()
    root.geometry("1500x600")
    root.title("Проведение изменения в задаче СЭЗ")

       
    panelFrame = Frame(root, height = 70, bg = 'grey')
    panelFrame.pack(side = 'top', fill = 'x')

    fieldFrame =Frame(root, height = 340, width = 900)
    fieldFrame.pack()

    '''
    name_fields=['Номер задачи', 'Заказчик', 'Куратор', 'Бизнес цель', 'Исполнитель',
                 'Имя устройства', 'Тестирование', 'config', 'Влиянеие на клиентов'
                 'Время отказа для клиентов', 'Обходные пути для клиентов', 'Влияние на пользователей',
                 'Время отказа для пользователей', 'Обходные пути для пользователей', 'Откат',
                 'Дата и время изменения', 'Продолжительность работ']

    width_entry =[5, 15, 15, 20, 15, 20, 15, 40, 20, 10, 15, 15, 15, 15, 20, 15, 15]
    '''
    global x
    global vr
    global rd2
    vr={}
    x=enumerate(dictionary)
    for r, elem in enumerate(dictionary):
        if elem=='config':

                yscroll=Scrollbar(fieldFrame, orient = VERTICAL)
    #            xscroll=Scrollbar(fieldFrame, orient = HORIZONTAL)
                

                rd1=Label(fieldFrame, text=dictionary[elem][0]+':').grid(row=7, column=2, sticky=W)
                vars()[str(elem)]=Text(fieldFrame, wrap=WORD, width = dictionary[elem][2], yscrollcommand = yscroll.set)
                vars()[str(elem)].insert(INSERT,dictionary[elem][1])
                vars()[str(elem)].grid(row=10, column=2, rowspan= 20, sticky=N+S)               
                yscroll.config( command = vars()[str(elem)].yview )           
  

                yscroll.grid(row=10, column=3, rowspan= 20)
                vr[elem]=vars()[str(elem)]

     
        else:
                rd3=Label(fieldFrame, text=dictionary[elem][0]+':').grid(row=r+10, column=0, sticky=E)
                #Генерим имена переменной поля Entry


#                elem=Entry(fieldFrame, width = dictionary[elem][2])
#                elem.insert(INSERT,dictionary[elem][1])
#                elem.grid(row=r+10, column=1, sticky=W)
                vars()[elem]=Entry(fieldFrame, width = dictionary[elem][2])
                vars()[elem].insert(INSERT,dictionary[elem][1])
                vars()[elem].grid(row=r+10, column=1, sticky=W)
#                vars()[str(elem)]=Entry(fieldFrame, width = dictionary[elem][2])
#                vars()[str(elem)].insert(INSERT,dictionary[elem][1])
#                vars()[str(elem)].grid(row=r+10, column=1, sticky=W)
                
                print(str(elem)+ '   - elem = {}'.format(vars()[str(elem)]))
                vr[elem]=vars()[elem]
#                print(vr)
                
#    print (vr)
                

    quitBtn = Button(panelFrame, text = 'Quit')
    quitBtn.place(x = 10, y = 10, width = 40, height = 40)
    quitBtn.bind("<Button-1>", Quit)

    writeBtn = Button(panelFrame, text = 'Записать изменение в шаблон')
    writeBtn.place(x = 60, y = 10, width = 200, height = 40)
    writeBtn.bind("<Button-1>", Write1)

    wordBtn = Button(panelFrame, text = 'Открыть файл изменений')
    wordBtn.place(x = 270, y = 10, width = 240, height = 40)
    wordBtn.bind("<Button-1>", Open_Word)

    izmBtn = Button(panelFrame, text = 'Отправить в СЭЗ на согласование')
    izmBtn.place(x = 540, y = 10, width = 240, height = 40)
    izmBtn.bind("<Button-1>", send_to_Sez)
    
    root.event_add('<<Paste>>', '<Control-igrave>')
    root.event_add("<<Copy>>", "<Control-ntilde>")
    root.mainloop()

