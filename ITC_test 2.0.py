from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import time

#WE NEED USE OPTION FOR THE SPECIAL GOOGLE ACC. CHECK chrome://version/. WITH SAVE PASS AND LOGIN
options = webdriver.ChromeOptions()
options.add_argument('--user-data-dir=C:/Users/dim-m/AppData/Local/Temp/scoped_dir14104_442793865/')
options.add_argument('--profile-directory=Profile 1')

#Choice driver
s = Service(r'C:\GitHub\ITC Upgrade Test\chromedriver_win32\chromedriver.exe')
driver = webdriver.Chrome(service=s, options=options)

#откроем для записей всех выгрузок
f = open('Errors.txt','w')

def Open_link(reporter):

    #open link
    driver.get('https://www.trademap.org/Bilateral_TS.aspx?nvpm=1%7c'+reporter+'%7c%7c%7c%7cTOTAL%7c%7c%7c2%7c1%7c1%7c2%7c1%7c1%7c1%7c1%7c1%7c1')

    #choose 6d cluster 5-6k of products

    for partner in range(254):
        try:
            #select partner
            select1 = Select(driver.find_element("name", "ctl00$NavigationControl$DropDownList_Partner"))
            select1.select_by_index(partner+1)

            #that workin only for autorizate accaounts!
            select3 = Select(driver.find_element("name", "ctl00$NavigationControl$DropDownList_ProductClusterLevel"))
            select3.select_by_visible_text("Product cluster at 6 digits")

            #choose 20 per-age
            select2 = Select(driver.find_element("name", "ctl00$PageContent$GridViewPanelControl$DropDownList_NumTimePeriod"))
            select2.select_by_visible_text("20 per page")

            #IDK how that working, but i think we need that...
            driver.implicitly_wait(20)

            #click to download
            element = driver.find_element("name", "ctl00$PageContent$GridViewPanelControl$ImageButton_ExportExcel")
            element.click()
            
            f.write('\n Выгрузка: ОК, partner ISO:'+str(partner))
        except: #если вылетели на мейн экран заново заходим и все настраиваем
            f.write('\n Выгрузка: except, partner ISO:'+str(partner))
            #open link
            driver.get('https://www.trademap.org/Bilateral_TS.aspx?nvpm=1%7c'+reporter+'%7c%7c%7c%7cTOTAL%7c%7c%7c2%7c1%7c1%7c2%7c2%7c1%7c1%7c1%7c1%7c1')

            #IDK how that working, but i think we need that...
            driver.implicitly_wait(10)

            #choose main reporter (sometimes that abyss)
            select4 = Select(driver.find_element("name", "ctl00$NavigationControl$DropDownList_Country"))
            select4.select_by_visible_text('Jordan')

            #choose 6d cluster 5-6k of products
            #that workin only for autorizate accaounts!
            select3 = Select(driver.find_element("name", "ctl00$NavigationControl$DropDownList_ProductClusterLevel"))
            select3.select_by_visible_text("Product cluster at 6 digits")

            #choose 20 per-age
            select2 = Select(driver.find_element("name", "ctl00$PageContent$GridViewPanelControl$DropDownList_NumTimePeriod"))
            select2.select_by_visible_text("20 per page")
            
            #и завново делаем выгрузку
            #select partner
            select1 = Select(driver.find_element("name", "ctl00$NavigationControl$DropDownList_Partner"))
            select1.select_by_index(partner+1)

            driver.implicitly_wait(20)

            #click to download
            element = driver.find_element("name", "ctl00$PageContent$GridViewPanelControl$ImageButton_ExportExcel")
            element.click()



#now that is all countre's
partArr = ['895', '036', '040', '031', '008', '012', '016', '660', '024', '020', '010', '028', '032', '051', '533', '004', '044', '050', '052', '048', '112', '084', '056', '204', '060', '100', '068', '535', '070', '072', '076', '086', '096', '854', '108', '064', '548', '348', '862', '092', '850', '704', '266', 
'332', '328', '270', '288', '312', '320', '324', '624', '276', '831', '292', '340', '344', '308', '304', '300', '268', '316', '208', '832', '262', '212', '214', '818', '894', '732', '716', '376', '356', '360', '400', '368', '364', '372', '352', '724', '380', '887', '132', '398', '116', '120', '124', '634', '404', '196', '417', '296', '156', '166', '170', '174', '178', '180', '408', '410', '188', '384', '192', '414', '531', '418', '428', '426', '422', '434', '430', '438', '440', '442', '480', '478', '450', '175', '446', '454', '458', '466', '581', '462', '470', '504', '474', '584', '484', '583', '508', '498', '492', '496', '500', '104', '516', '520', '524', '562', '566', '528', '558', '570', '554', '540', '578', '784', '512', '074', '833', '574', '162', '334', '136', '184', '796', '586', '585', '275', '591', '336', '598', '600', '604', '612', '616', '620', '630', '807', '638', '643', '646', '642', '882', '674', '678', '682', '654', '580', '652', '663', '686', '670', '659', '662', '666', '688', '690', '702', '534', '760', '703', '705', '826', 
'840', '090', '706', '729', '740', '694', '762', '764', '158', '834', '626', '768', '772', '776', '780', '798', '788', '795', '792', '800', '860', '804', '876', '858', '234', '242', '608', '246', '238', '250', '254', '258', '260', '191', '140', '148', '499', '203', '152', '756', '752', '744', '144', '218', '226', '248', '222', '232', '748', '233', '231', '710', '239', '896', '728', '388', '392'] #thats will be array of all pathers))

reporters = ['376', '792', '400','076','032','458','360','504','710','764','288','496','788','784','586','608','566','634','414','512','116','682','422','858','152']

#for reporter in reporters:
#    Open_link(reporter)

Open_link('400')
f.close()

        
        

#partner = '156'
#Open_link(reporter, partner)






#good RU guide - https://habr.com/en/company/otus/blog/596071/