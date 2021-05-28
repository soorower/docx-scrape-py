import os
from time import sleep
import pandas as pd
from bs4 import BeautifulSoup as bs
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from shutil import which

chrome_options = Options()
prefs = {'download.default_directory' : os.getcwd()}
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument("--headless")

chrome_path = which('chromedriver')
driver = webdriver.Chrome(executable_path=chrome_path, options= chrome_options)


#----------------------------------Agenda----------------------------------------------------------------------
def scrape_agenda():
    link = input('Provide doc link: ')
    link_name = link.split('/')[-1]
    link_test = link.split('/')[-1].split('.')[0]
    cks = []
    try:
        for file in os.listdir():
            if f'{link_test}' in file:
                cks.append(1)
                if len(cks)==1:
                    print('Uploading doc...')
                    driver.get('https://products.groupdocs.app/editor/docx')
                    driver.maximize_window()
                    sleep(1)
                    driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
                    print('Scraping Starts...')
                    sleep(20)
        s = cks[-1]
    
    except:
        print('Downloading doc...')
        driver.get(link)
        driver.maximize_window()
        sleep(3)
        print('Uploading doc...')
        driver.get('https://products.groupdocs.app/editor/docx')
        sleep(1)
        driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
        print('Scraping Starts...')
        print('\n')
        sleep(20)
    lists = []
    data = {}
    soup = bs(driver.page_source, 'html.parser')
    table_find = soup.find('div',attrs = {'class':'Section1'}).findAll('table')
    for table in table_find:
        if 'Listing requested by Sponsor / Purpose of Submission' in table.get_text():
            tbody = table.find('tbody').findAll('tr')
            for tr in tbody:
                all_tds = tr.findAll('td')
                listing = all_tds[0].get_text()

                drug_name = all_tds[1].findAll('p')[0].get_text().strip().title()
                sponsor = all_tds[1].findAll('p')[-1].get_text()  
                trade_name = all_tds[1].findAll('p')[-3].get_text()
                if trade_name==drug_name:
                    trade_name = ' '
                forms = all_tds[1].findAll('p')[1:-3]
                form = ''
                for fom in forms:
                    form = form + '\n' + fom.get_text()

                requested_listing = all_tds[3].get_text()
                drug_uses = all_tds[2].findAll('p')
                drug_use = ''
                for drug in drug_uses:
                    drug_use = drug_use + '\n' + drug.get_text()
                data = {
                    'Document Type': 'Agenda',
                    'Listing':listing.strip(),
                    'Drug Name': drug_name,
                    'Form': form.strip(),
                    'Trade Name': trade_name.strip(),
                    'Sponsor': sponsor.strip(),
                    'Drug Use': drug_use.strip(),
                    'Requested Listing': requested_listing.strip()
                }
                lists.append(data)
    df = pd.DataFrame(lists).reset_index(drop=True)
    df.to_csv(f'{link_test}.csv',encoding = 'utf-8-sig',index=False)
    print(f'''Done...Look for a csv file : {link_test}
location: {os.getcwd()}''')
    
    
#-------------------Positive Recommendation-------------------------------------------------------      
def positive_recom():
    link = input('Provide doc link: ')
    link_name = link.split('/')[-1]
    link_test = link.split('/')[-1].split('.')[0]
    cks = []
    try:
        for file in os.listdir():
            if f'{link_test}' in file:
                cks.append(1)
                if len(cks)==1:
                    print('Uploading doc...')
                    driver.get('https://products.groupdocs.app/editor/docx')
                    driver.maximize_window()
                    sleep(1)
                    driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
                    print('Scraping Starts...')
                    sleep(20)
        s = cks[-1]
    
    except:
        print('Downloading doc...')
        driver.get(link)
        driver.maximize_window()
        sleep(3)
        print('Uploading doc...')
        driver.get('https://products.groupdocs.app/editor/docx')
        sleep(1)
        driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
        print('Scraping Starts...')
        print('\n')
        sleep(20)

    lists = []
    data = {}
    soup = bs(driver.page_source, 'html.parser')
    table = soup.find('div',attrs = {'class':'Section1'}).find('table')
    tbody = table.find('tbody').findAll('tr')

    for tr in tbody:
        all_tds = tr.findAll('td')
        drug_name = all_tds[0].find('p').find('span').get_text()
        is_drug = all_tds[0].find('p').findAll('span')[1].get_text()
        if is_drug.isupper():
            drug_name = drug_name + '+'+'\n' + is_drug
        drug_name = drug_name.strip().title()
    #-------------------------------------------------------------------------
        trade_box = []
        check_p = all_tds[0].findAll('p')
        trade_name = ''
        if len(check_p)==1:
            trade_check = all_tds[0].find('p').findAll('span')
            for td_check in trade_check:
                if '®' in td_check.get_text():
                    trade_name =  td_check.get_text()+ '\n' + trade_name
                    trade_box.append(td_check.get_text())
        else:
            trade_check = all_tds[0].findAll('p')[1].findAll('span')
            for td_check in trade_check:
                if '®' in td_check.get_text():
                    trade_name =  td_check.get_text() + '\n' + trade_name
                    trade_box.append(td_check.get_text())
    #--------------------------------------------------------------------------                
        if len(check_p)==1:
            all_text = all_tds[0].find('p').findAll('span')
            for al_text in all_text:
                if 'Submission' in al_text.text:
                    type_submission = al_text.text
                if 'submission' in al_text.text:
                    type_submission = al_text.text
        else:
            all_text = all_tds[0].findAll('p')[1].findAll('span')
            for al_text in all_text:
                if 'Submission' in al_text.text:
                    type_submission = al_text.text
                if 'submission' in al_text.text:
                    type_submission = al_text.text
    #----------------------------------------------------------------------------
        if len(check_p)==1:
            all_text = all_tds[0].find('p').findAll('span')
            for al_text in all_text:
                if 'listing' in al_text.text:
                    listing = al_text.text
                if 'Listing' in al_text.text:
                    listing = al_text.text
        else:
            all_text = all_tds[0].findAll('p')[1].findAll('span')
            for al_text in all_text:
                if 'listing' in al_text.text:
                    listing = al_text.text
                if 'Listing' in al_text.text:
                    listing = al_text.text
    #-------------------------------------------------------------------------------    
        sponso = []
        try:
            if len(check_p)==1:
                all_text = all_tds[0].find('p').findAll('span')
                for al_text in all_text:
                    if 'limit' in al_text.text:
                        sponsor = al_text.text
                        sponso.append(sponsor)
                    if 'Limit' in al_text.text:
                        sponsor = al_text.text
                        sponso.append(sponsor)
                    if 'ltd' in al_text.text:
                        sponsor = al_text.text
                        sponso.append(sponsor)
                    if 'Ltd' in al_text.text:
                        sponsor = al_text.text
                sponso.append(sponsor)
            else:
                all_text = all_tds[0].findAll('p')[1].findAll('span')
                for al_text in all_text:
                    if 'limit' in al_text.text:
                        sponsor = al_text.text
                        sponso.append(sponsor)
                    if 'Limit' in al_text.text:
                        sponsor = al_text.text
                        sponso.append(sponsor)
                    if 'ltd' in al_text.text:
                        sponsor = al_text.text
                        sponso.append(sponsor)
                    if 'Ltd' in al_text.text:
                        sponsor = al_text.text
                sponso.append(sponsor)

            sponsor = sponso[-1]
        except:
            if len(check_p)==1:
                sponsor = all_tds[0].find('p').findAll('span')[-3].get_text()
                x = all_tds[0].find('p').findAll('span')[-1].get_text()
                if u'\xa0' in sponsor:
                    sponsor  = all_tds[0].find('p').findAll('span')[-4].get_text()
                if u'\xa0' in x:
                    sponsor  = all_tds[0].find('p').findAll('span')[-4].get_text()
            else:
                sponsor = all_tds[0].findAll('p')[1].findAll('span')[-3].get_text()
    #--------------------------------------------------------------------------------
        texts = ''
        if len(check_p)==1:
            for sp in all_tds[0].find('p').findAll('span'):
                texts = texts +'\n'+ sp.text
        else:
            for sp in all_tds[0].find('p').findAll('span'):
                texts = texts +'\n'+ sp.text
            for sp in all_tds[0].findAll('p')[1].findAll('span'):
                texts = texts +'\n'+ sp.text
        forms=[texts]
        for i in trade_box:
            form = forms[-1].replace(i,'').replace(drug_name,'').replace(sponsor,'').replace(listing,'').replace(type_submission,'').replace(type_submission,'')
            if is_drug.isupper():
                form = form.replace(is_drug,'')
            forms.append(form)
        form = form.strip()
    #----------------------------------------------------------------------------------
        drug_use = all_tds[1].get_text()
        requested_listing = all_tds[2].get_text()
        outcome = all_tds[3].get_text()
        data = {
            'Type': 'Positive Recommendations',
            'Drug Name': drug_name,
            'Form':form.strip(),
            'Trade Name': trade_name.strip(),
            'Sponsor': sponsor.strip(),
            'Listing': listing.strip(),
            'Type Submission': type_submission.strip(),
            'Drug Use': drug_use.strip(),
            'Requested Listing': requested_listing.strip(),
            'Outcome':outcome.strip()
        }
        lists.append(data)
    df = pd.DataFrame(lists).reset_index(drop=True)
    df.to_csv(f'{link_test}.csv',encoding = 'utf-8-sig',index=False)
    print(f'''Done...Look for a csv file : {link_test}
location: {os.getcwd()}''')
    
#-----------------------------------------------------------------------------------
#---------------------------------First Time Decisions-------------------------------------------------------------------
def first_time_decisions():
    link = input('Provide doc link: ')
    link_name = link.split('/')[-1]
    link_test = link.split('/')[-1].split('.')[0]
    cks = []
    try:
        for file in os.listdir():
            if f'{link_test}' in file:
                cks.append(1)
                if len(cks)==1:
                    print('Uploading doc...')
                    driver.get('https://products.groupdocs.app/editor/docx')
                    driver.maximize_window()
                    sleep(1)
                    driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
                    print('Scraping Starts...')
                    sleep(20)
        s = cks[-1]
    
    except:
        print('Downloading doc...')
        driver.get(link)
        driver.maximize_window()
        sleep(3)
        print('Uploading doc...')
        driver.get('https://products.groupdocs.app/editor/docx')
        sleep(1)
        driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
        print('Scraping Starts...')
        print('\n')
        sleep(20)

    soup = bs(driver.page_source, 'html.parser')
    table = soup.find('div',attrs = {'class':'Section1'}).find('table')
    tbody = table.find('tbody').findAll('tr')
    counter = 1
    drug_nam = []
    trade_nam = []
    type_submissio = []
    listin = []
    sponso = []
    forme = []
    drug_us = []
    requested_listin = []
    outcom = []
    sponsor_commen = []

    lists = []
    data = {}
    for tr in tbody:
        all_tds = tr.findAll('td')

        if 'Sponsor’s comment' in tr.get_text():
            pass
        elif 'sponsor’s comment' in tr.get_text():
            pass
        elif 'sponsor comment' in tr.get_text():
            pass
        elif 'Sponsor’s Comment' in tr.get_text():
            pass
        else:
            drug_name = all_tds[0].find('p').find('span').get_text().strip().title()
            drug_nam.append(drug_name)
        #---------------------------------------------------------------------------
            trade_name = ''
            check_p = all_tds[0].findAll('p')
            if len(check_p)>1:
                for ch_p in check_p:
                    if '®' in ch_p.find('span').get_text():
                        trade_name =  ch_p.find('span').get_text()+ '\n' + trade_name
            else:
                trade_check = all_tds[0].find('p').findAll('span')
                for td_check in trade_check:
                    if '®' in td_check.get_text():
                        trade_name =  td_check.get_text()+ '\n' + trade_name
            trade_name = trade_name.strip()
            trade_nam.append(trade_name)
        #--------------------------------------------------------------------------
            if len(check_p)>1:
                for ch_p in check_p:
                    if 'Submission' in ch_p.find('span').get_text():
                        type_submission = ch_p.find('span').get_text()
                    if 'submission' in ch_p.find('span').get_text():
                        type_submission = ch_p.find('span').get_text()
            else:
                all_text = all_tds[0].find('p').findAll('span')
                for al_text in all_text:
                    if 'Submission' in al_text.text:
                        type_submission = al_text.text
                    if 'submission' in al_text.text:
                        type_submission = al_text.text
            type_submissio.append(type_submission)
        #--------------------------------------------------------------------------
            if len(check_p)>1:
                for ch_p in check_p:
                    if 'Listing' in ch_p.find('span').get_text():
                        listing = ch_p.find('span').get_text()
                    if 'listing' in ch_p.find('span').get_text():
                        listing = ch_p.find('span').get_text()
            else:
                all_text = all_tds[0].find('p').findAll('span')
    #             print(all_text)
                for al_text in all_text:
                    if 'Listing' in al_text.text:
                        listing = al_text.text
                    if 'listing' in al_text.text:
                        listing = al_text.text
            listin.append(listing)
    #-----------------------------------------------------------------
   
            for ch_p in check_p:
                for al_sp in ch_p.findAll('span'):
                    if 'Ltd' in al_sp.get_text():
                        sponsor = al_sp.get_text()
                    if 'ltd' in al_sp.get_text():
                        sponsor = al_sp.get_text()
                    if 'limit' in al_sp.get_text():
                        sponsor = al_sp.get_text()
                    if 'Limit' in al_sp.get_text():
                        sponsor = al_sp.get_text()
            sponso.append(sponsor)

        ##------------------------------------------------------------------------------
            form_text = ''
            if len(check_p)==1:
                texts = all_tds[0].find('p').findAll('span')
                for text in texts:
                    form_text = form_text +'\n' + text.get_text() 
            elif len(check_p)==2:
                texts = all_tds[0].findAll('p')
                for text in texts:
                    for each_sp in text.findAll('span'):
                        form_text = form_text + '\n' + each_sp.get_text()
            else:
                texts = all_tds[0].findAll('p')
                for text in texts:
                    form_text = form_text + '\n' + text.find('span').get_text()
            form = form_text.replace(trade_name,'').replace(drug_name,'').replace(sponsor,'').replace(listing,'').replace(type_submission,'').replace(type_submission,'').strip()
            forme.append(form)
            drug_use = all_tds[1].get_text()
            drug_us.append(drug_use)
            requested_listing = all_tds[2].get_text()
            requested_listin.append(requested_listing)
            outcome = all_tds[3].get_text()
            outcom.append(outcome)
        try:  
            if 'Sponsor’s comment' in tr.get_text():
                try:
                    sponsor_comment = tr.findAll('td')[-1].get_text()
                except:
                    sponsor_comment = '-'
                sponsor_commen.append(sponsor_comment)
        except:
            pass
        if counter%2==0:
            data = {
                'Drug Name': drug_nam[-1],
                'Form': forme[-1].strip(),
                'Trade Name': trade_nam[-1].strip(),
                'Sponsor': sponso[-1].strip(),
                'Listing': listin[-1].strip(),
                'Type Submission': type_submissio[-1].strip(),
                'Drug Use': drug_us[-1].strip(),
                'Requested Listing': requested_listin[-1].strip(),
                'Outcome':outcom[-1].strip(),
                'Sponsor Comment': sponsor_commen[-1].strip()
            }
            lists.append(data)
        counter = counter + 1
    df = pd.DataFrame(lists).reset_index(drop=True)
    df.to_csv(f'{link_test}.csv',encoding = 'utf-8-sig',index=False)
    print(f'''Done...Look for a csv file : {link_test}
location: {os.getcwd()}''')
    
    
#-----------------------------subsequent decisions------------------------------------------------------------------
def subsequent_decision():
    link = input('Provide doc link: ')
    link_name = link.split('/')[-1]
    link_test = link.split('/')[-1].split('.')[0]
    cks = []
    try:
        for file in os.listdir():
            if f'{link_test}' in file:
                cks.append(1)
                if len(cks)==1:
                    print('Uploading doc...')
                    driver.get('https://products.groupdocs.app/editor/docx')
                    driver.maximize_window()
                    sleep(1)
                    driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
                    print('Scraping Starts...')
                    sleep(20)
        s = cks[-1]
    
    except:
        print('Downloading doc...')
        driver.get(link)
        driver.maximize_window()
        sleep(3)
        print('Uploading doc...')
        driver.get('https://products.groupdocs.app/editor/docx')
        sleep(1)
        driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
        print('Scraping Starts...')
        print('\n')
        sleep(20)

    soup = bs(driver.page_source, 'html.parser')
    table = soup.find('div',attrs = {'class':'Section1'}).find('table')
    tbody = table.find('tbody').findAll('tr')
    counter = 1
    drug_nam = []
    trade_nam = []
    type_submissio = []
    listin = []
    sponso = []
    forme = []
    drug_us = []
    tga_indicatio = []
    current_listin = []
    requested_listin = []
    outcom = []
    comparato = []
    comparator_outcom = []
    clinical_clai = []
    clinical_claim_outcom = []
    economic_claim_outcom = []
    economic_clai = []
    sponsor_commen = []

    lists = []
    data = {}
    for tr in tbody:
        all_tds = tr.findAll('td')

        if len(all_tds)>2:
            drug_name = all_tds[0].find('p').find('span').get_text().strip().title()
            drug_nam.append(drug_name)
    #------------------------------------------------------------------------------------
            trade_name = ''
            check_p = all_tds[0].findAll('p')
            if len(check_p)>1:
                for ch_p in check_p:
                    if '®' in ch_p.find('span').get_text():
                        trade_name =  ch_p.find('span').get_text()+ '\n' + trade_name
            else:
                trade_check = all_tds[0].find('p').findAll('span')
                for td_check in trade_check:
                    if '®' in td_check.get_text():
                        trade_name =  td_check.get_text()+ '\n' + trade_name
            trade_name = trade_name.strip()
            trade_nam.append(trade_name)
    #-------------------------------------------------------------------------------------
            if len(check_p)>1:
                for ch_p in check_p:
                    if 'Submission' in ch_p.find('span').get_text():
                        type_submission = ch_p.find('span').get_text()
                    if 'submission' in ch_p.find('span').get_text():
                        type_submission = ch_p.find('span').get_text()
            else:
                all_text = all_tds[0].find('p').findAll('span')
                for al_text in all_text:
                    if 'Submission' in al_text.text:
                        type_submission = al_text.text
                    if 'submission' in al_text.text:
                        type_submission = al_text.text
            type_submissio.append(type_submission)
    #-------------------------------------------------------------------------------------
            if len(check_p)>1:
                for ch_p in check_p:
                    if 'Listing' in ch_p.find('span').get_text():
                        listing = ch_p.find('span').get_text()
                    if 'listing' in ch_p.find('span').get_text():
                        listing = ch_p.find('span').get_text()
            else:
                all_text = all_tds[0].find('p').findAll('span')
                for al_text in all_text:
                    if 'Listing' in al_text.text:
                        listing = al_text.text
                    if 'listing' in al_text.text:
                        listing = al_text.text
            listin.append(listing)
    #----------------------------------------------------------------------------------------
            if len(check_p)>1:
                for ch_p in check_p:
                    for al_sp in ch_p.findAll('span'):
                        if 'Ltd' in al_sp.get_text():
                            sponsor = al_sp.get_text()
                        if 'ltd' in al_sp.get_text():
                            sponsor = al_sp.get_text()
                        if 'limit' in al_sp.get_text():
                            sponsor = al_sp.get_text()
                        if 'Limit' in al_sp.get_text():
                            sponsor = al_sp.get_text()
            else:
                all_text = all_tds[0].find('p').findAll('span')
                for al_text in all_text:
                    if 'Ltd' in al_text.text:
                        sponsor = al_text.text
                    if 'ltd' in al_text.text:
                        sponsor = al_text.text
                    if 'Limit' in al_text.text:
                        sponsor = al_text.text
                    if 'limit' in al_text.text:
                        sponsor = al_text.text
            sponso.append(sponsor)
    #----------------------------------------------------------------------------------
            form_text = ''
            if len(check_p)==1:
                texts = all_tds[0].find('p').findAll('span')
                for text in texts:
                    form_text = form_text +'\n' + text.get_text() 
            elif len(check_p)==2:
                texts = all_tds[0].findAll('p')
                for text in texts:
                    for each_sp in text.findAll('span'):
                        form_text = form_text + '\n' + each_sp.get_text()
            else:
                texts = all_tds[0].findAll('p')
                for text in texts:
                    form_text = form_text + '\n' + text.find('span').get_text()
            form = form_text.replace(trade_name,'').replace(drug_name,'').replace(sponsor,'').replace(listing,'').replace(type_submission,'').replace(type_submission,'').strip()
            forme.append(form)
            tga_indication = all_tds[1].get_text()
            tga_indicatio.append(tga_indication)
            current_listing = all_tds[2].get_text()
            current_listin.append(current_listing)
            requested_listing = all_tds[3].get_text()
            requested_listin.append(requested_listing)
            outcome = all_tds[4].get_text()
            outcom.append(outcome)
        else:
            pass
        try:  
            if len(all_tds)==2:
                if 'Comparator' in tr.findAll('td')[0].get_text():
                    try:
                        comparator = tr.findAll('td')[0].get_text().replace('Comparator:','')
                    except:
                        comparator = '-'
                    comparato.append(comparator.strip())
                    try:
                        comparator_outcome = tr.findAll('td')[1].get_text()
                    except:
                        comparator_outcome = '-'
                    comparator_outcom.append(comparator_outcome)
                if 'Clinical claim' in tr.findAll('td')[0].get_text():
                    try:
                        clinical_claim = tr.findAll('td')[0].get_text().replace('Clinical claim:','')
                    except:
                        clinical_claim = '-'
                    clinical_clai.append(clinical_claim.strip())
                    try:
                        clinical_claim_outcome = tr.findAll('td')[1].get_text()
                    except:
                        clinical_claim_outcome = '-'
                    clinical_claim_outcom.append(clinical_claim_outcome)
                if 'Economic claim' in tr.findAll('td')[0].get_text():
                    try:
                        economic_claim = tr.findAll('td')[0].get_text().replace('Economic claim:','')
                    except:
                        economic_claim = '-'
                    economic_clai.append(economic_claim.strip())
                    try:
                        economic_claim_outcome = tr.findAll('td')[1].get_text()
                    except:
                        economic_claim_outcome = '-'
                    economic_claim_outcom.append(economic_claim_outcome)
                if 'Sponsor’s comment' in tr.findAll('td')[0].get_text():
                    try:
                        sponsor_comment = tr.findAll('td')[1].get_text()
                    except:
                        sponsor_comment = '-'
                    sponsor_commen.append(sponsor_comment)
        except:
            pass
        if counter%5==0:
            data = {
                'Type': 'Subsequent Decisions',
                'Drug Name': drug_nam[-1],
                'Form': forme[-1].strip(),
                'Trade Name': trade_nam[-1].strip(), 
                'Sponsor': sponso[-1].strip(),
                'Listing': listin[-1].strip(),
                'Type Submission': type_submissio[-1].strip(),
                'TGA Indication': tga_indicatio[-1].strip(),
                'Current Listing': current_listin[-1].strip(),
                'Requested Listing': requested_listin[-1].strip(),
                'Outcome': outcom[-1].strip(),
                'Comparator': comparato[-1].strip(),
                'Comparator Outcome': comparator_outcom[-1].strip(),
                'Clinical Claim': clinical_clai[-1].strip(),
                'Clinical Claim Outcome': clinical_claim_outcom[-1].strip(),
                'Economic Claim': economic_clai[-1].strip(),
                'Economic Claim Outcome': economic_claim_outcom[-1].strip(),
                'Sponsor’s comment': sponsor_commen[-1].strip()
            }
            lists.append(data)
        counter = counter + 1
    df = pd.DataFrame(lists).reset_index(drop=True)
    df.to_csv(f'{link_test}.csv',encoding = 'utf-8-sig',index=False)
    print(f'''Done...Look for a csv file : {link_test}
location: {os.getcwd()}''')

    
#----------------------------deferral scrape------------------------------------------------------------------
def scrape_deferral():
    link = input('Provide doc link: ')
    link_name = link.split('/')[-1]
    link_test = link.split('/')[-1].split('.')[0]
    cks = []
    try:
        for file in os.listdir():
            if f'{link_test}' in file:
                cks.append(1)
                if len(cks)==1:
                    print('Uploading doc...')
                    driver.get('https://products.groupdocs.app/editor/docx')
                    driver.maximize_window()
                    sleep(1)
                    driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
                    print('Scraping Starts...')
                    sleep(20)
        s = cks[-1]
    
    except:
        print('Downloading doc...')
        driver.get(link)
        driver.maximize_window()
        sleep(3)
        print('Uploading doc...')
        driver.get('https://products.groupdocs.app/editor/docx')
        sleep(1)
        driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
        print('Scraping Starts...')
        print('\n')
        sleep(20)

    soup = bs(driver.page_source, 'html.parser')
    table = soup.find('div',attrs = {'class':'Section1'}).find('table')
    tbody = table.find('tbody').findAll('tr')
    counter = 1
    drug_nam = []
    trade_nam = []
    type_submissio = []
    listin = []
    sponso = []
    forme = []
    drug_us = []
    requested_listin = []
    outcom = []
    sponsor_commen = []

    lists = []
    data = {}
    for tr in tbody:
        all_tds = tr.findAll('td')

        if 'Sponsor’s comment' in tr.get_text():
            pass
        elif 'sponsor’s comment' in tr.get_text():
            pass
        elif 'sponsor comment' in tr.get_text():
            pass
        elif 'Sponsor’s Comment' in tr.get_text():
            pass
        else:
            drug_name1 = all_tds[0].find('p').find('span').get_text().strip().title()
            drug_nam.append(drug_name1)
            
        #---------------------------------------------------------------------------
            trade_name = ''
            check_p = all_tds[0].findAll('p')
            if len(check_p)>1:
                for ch_p in check_p:
                    for each_sp in ch_p.findAll('span'):
                        if '®' in each_sp.get_text():
                            trade_name =  each_sp.get_text()+ '\n' + trade_name
            else:
                trade_check = all_tds[0].find('p').findAll('span')
                for td_check in trade_check:
                    if '®' in td_check.get_text():
                        trade_name =  td_check.get_text()+ '\n' + trade_name
            trade_name = trade_name.strip()
            trade_nam.append(trade_name)
        #--------------------------------------------------------------------------
            if len(check_p)>1:
                for ch_p in check_p:
                    for each_sp in ch_p.findAll('span'):
                        if 'Submission' in each_sp.get_text():
                            type_submission = each_sp.get_text()
                        if 'submission' in each_sp.get_text():
                            type_submission = each_sp.get_text()
            else:
                all_text = all_tds[0].find('p').findAll('span')
                for al_text in all_text:
                    if 'Submission' in al_text.text:
                        type_submission = al_text.text
                    if 'submission' in al_text.text:
                        type_submission = al_text.text
            type_submissio.append(type_submission)
        #--------------------------------------------------------------------------
            if len(check_p)>1:
                for ch_p in check_p:
                    for each_sp in ch_p.findAll('span'):
                        if 'Listing' in each_sp.get_text():
                            listing = each_sp.get_text()
                        if 'listing' in each_sp.get_text():
                            listing = each_sp.get_text()
            else:
                all_text = all_tds[0].find('p').findAll('span')
                for al_text in all_text:
                    if 'Listing' in al_text.text:
                        listing = al_text.text
                    if 'listing' in al_text.text:
                        listing = al_text.text
            listin.append(listing)
        #-----------------------------------------------------------------
            if len(check_p)>1:
                for ch_p in check_p:
                    for al_sp in ch_p.findAll('span'):
                        if 'Ltd' in al_sp.get_text():
                            sponsor = al_sp.get_text()
                        if 'ltd' in al_sp.get_text():
                            sponsor = al_sp.get_text()
                        if 'limit' in al_sp.get_text():
                            sponsor = al_sp.get_text()
                        if 'Limit' in al_sp.get_text():
                            sponsor = al_sp.get_text()
            else:
                all_text = all_tds[0].find('p').findAll('span')
                for al_text in all_text:
                    if 'Ltd' in al_text.text:
                        sponsor = al_text.text
                    if 'ltd' in al_text.text:
                        sponsor = al_text.text
                    if 'Limit' in al_text.text:
                        sponsor = al_text.text
                    if 'limit' in al_text.text:
                        sponsor = al_text.text
            sponso.append(sponsor)
        #         
        ##------------------------------------------------------------------------------
            form_text = ''
            texts = all_tds[0].findAll('p')
            for text in texts:
                for each_sp in text.findAll('span'):
                    form_text = form_text + '\n' + each_sp.get_text()
            form = form_text.replace(trade_name,'').replace(drug_name1,'').replace(sponsor,'').replace(listing,'').replace(type_submission,'').strip()
            forme.append(form)
            drug_use = all_tds[1].get_text()
            drug_us.append(drug_use)
            requested_listing = all_tds[2].get_text()
            requested_listin.append(requested_listing)
            outcome = all_tds[3].get_text()
            outcom.append(outcome)


        try:  
            if 'Sponsor’s comment' in tr.get_text():
                try:
                    sponsor_comment = tr.findAll('td')[-1].get_text()
                except:
                    sponsor_comment = '-'
                sponsor_commen.append(sponsor_comment)
        except:
            pass

        if counter%2==0:
            data = {
                'Type': 'Deferrals',
                'Drug Name': drug_nam[-1],
                'Form': forme[-1].strip(),
                'Trade Name': trade_nam[-1].strip(),
                'Sponsor':sponso[-1].strip(),
                'Listing':listin[-1].strip(),
                'Type Submission': type_submissio[-1].strip(),
                'Drug Use': drug_us[-1].strip(),
                'Requsted Listing': requested_listin[-1].strip(),
                'Outcome': outcom[-1].strip(),
                'Sponsor’s comment': sponsor_commen[-1].strip()
            }
            lists.append(data)
        counter = counter + 1
   
    df = pd.DataFrame(lists).reset_index(drop=True)
    df.to_csv(f'{link_test}.csv',encoding = 'utf-8-sig',index=False)
    print(f'''Done...Look for a csv file : {link_test}
location: {os.getcwd()}''')


def other_matters():
    link = input('Provide doc link: ')
    link_name = link.split('/')[-1]
    link_test = link.split('/')[-1].split('.')[0]
    cks = []
    try:
        for file in os.listdir():
            if f'{link_test}' in file:
                cks.append(1)
                if len(cks)==1:
                    print('Uploading doc...')
                    driver.get('https://products.groupdocs.app/editor/docx')
                    driver.maximize_window()
                    sleep(1)
                    driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
                    print('Scraping Starts...')
                    sleep(20)
        s = cks[-1]
    
    except:
        print('Downloading doc...')
        driver.get(link)
        driver.maximize_window()
        sleep(3)
        print('Uploading doc...')
        driver.get('https://products.groupdocs.app/editor/docx')
        sleep(1)
        driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
        print('Scraping Starts...')
        print('\n')
        sleep(20)

    soup = bs(driver.page_source, 'html.parser')
    lists = []
    data = {}
    tables = soup.find('div',attrs = {'class':'Section1'}).findAll('table')
    all_trs = []
    for table in tables:
        try:
            for tr in table.find('tbody').findAll('tr'):
                if 'drug name' in tr.get_text().lower():
                    pass
                else:
                    all_trs.append(tr)
        except:
            for tr in table.find('thead').findAll('tr'):
                if 'drug name' in tr.get_text().lower():
                    pass
                else:
                    all_trs.append(tr)


    for tr in all_trs:
        all_tds = tr.findAll('td')
        drug_name = all_tds[0].find('p').find('span').get_text().strip().title()

    #-----------------------------------------------------------------------
        trade_name = ''
        check_p = all_tds[0].findAll('p')
        if len(check_p)>1:
            for ch_p in check_p:
                for each_sp in ch_p.findAll('span'):
                    if '®' in each_sp.get_text():
                        trade_name =  each_sp.get_text()+ '\n' + trade_name
        else:
            trade_check = all_tds[0].find('p').findAll('span')
            for td_check in trade_check:
                if '®' in td_check.get_text():
                    trade_name =  td_check.get_text()+ '\n' + trade_name

        try:
            trade_name = trade_name.strip()
        except:
            trade_name = ' '
    #---------------------------------------------------------------
        type_sub = ''
        for ch_p in check_p:
            for each_sp in ch_p.findAll('span'):
                if '(Other)' in each_sp.get_text():
                    type_submission = each_sp.get_text()
                    type_sub = type_submission + type_sub
                if 'submission' in each_sp.get_text().lower():
                    type_submission = each_sp.get_text()
                    type_sub = type_submission + type_sub
                if '(other)' in each_sp.get_text():
                    type_submission = each_sp.get_text()
                    type_sub = type_submission + type_sub
        type_submission = type_sub
    #     print(type_submission)
    #---------------------------------------------------------------------
        listin = ''
        for ch_p in check_p:
            for each_sp in ch_p.findAll('span'):
                if 'listing' in each_sp.get_text().lower():
                    listing = each_sp.get_text()
                    listin = listing + listin
                if 'listed' in each_sp.get_text().lower():
                    listing = each_sp.get_text()
                    listin = listing + listin
        listing = listin
    #     print(listing)

        sponso = ''
        for ch_p in check_p:
            for al_sp in ch_p.findAll('span'):
                if 'Ltd' in al_sp.get_text():
                    sponsor = al_sp.get_text()
                    sponso = sponsor + sponso
                if 'ltd' in al_sp.get_text():
                    sponsor = al_sp.get_text()
                    sponso = sponsor + sponso
                if 'limit' in al_sp.get_text():
                    sponsor = al_sp.get_text()
                    sponso = sponsor + sponso
                if 'Limit' in al_sp.get_text():
                    sponsor = al_sp.get_text()
                    sponso = sponsor + sponso
        sponsor = sponso
    #-------------------------------------------------------------------------------    
        form_text = ''
        texts = all_tds[0].findAll('p')
        for text in texts:
            for each_sp in text.findAll('span'):
                form_text = form_text + '\n' + each_sp.get_text()
        form = form_text.replace(trade_name,'').replace(drug_name,'').replace(sponsor,'').replace(listing,'').replace(type_submission,'').strip()

        drug_use = all_tds[1].get_text()
        requested_listing = all_tds[2].get_text()
        outcome = all_tds[3].get_text()
        data = {
            'Type': 'Other Matters',
            'Drug Name': drug_name,
            'Form': form.strip(),
            'Trade Name': trade_name.strip(),
            'Sponsor':sponsor.strip(),
            'Listing':listing.strip(),
            'Type Submission': type_submission.strip(),
            'Drug Use': drug_use.strip(),
            'Requsted Listing': requested_listing.strip(),
            'Outcome': outcome.strip()
        }
        lists.append(data)
    df = pd.DataFrame(lists).reset_index(drop=True)
    df.to_csv(f'{link_test}.csv',encoding = 'utf-8',index=False)
    print(f'''Done...Look for a csv file : {link_test}
location: {os.getcwd()}''')


def web_outcomes():
    link = input('Provide doc link: ')
    link_name = link.split('/')[-1]
    link_test = link.split('/')[-1].split('.')[0]
    cks = []
    try:
        for file in os.listdir():
            if f'{link_test}' in file:
                cks.append(1)
                if len(cks)==1:
                    print('Uploading doc...')
                    driver.get('https://products.groupdocs.app/editor/docx')
                    driver.maximize_window()
                    sleep(1)
                    driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
                    print('Scraping Starts...')
                    sleep(20)
        s = cks[-1]
    
    except:
        print('Downloading doc...')
        driver.get(link)
        driver.maximize_window()
        sleep(3)
        print('Uploading doc...')
        driver.get('https://products.groupdocs.app/editor/docx')
        sleep(1)
        driver.find_element_by_xpath("//input[@class='uploadfileinput']").send_keys(os.getcwd()+'/'+link_name)
        print('Scraping Starts...')
        print('\n')
        sleep(20)

    drug_nam = []
    trade_nam = []
    type_submissio = []
    listins = []
    sponsos = []
    forme = []
    drug_us = []
    requested_listin = []
    pbac_reco = []
    pbac_outcom = []
    sponsor_commen = []  



    soup = bs(driver.page_source, 'html.parser')
    lists = []
    data = {}
    tables = soup.find('div',attrs = {'class':'Section1'}).findAll('table')
    all_trs = []
    for table in tables:
        if 'drug type and use' in table.get_text().lower():
            try:
                for tr in table.find('tbody').findAll('tr'):
                    if 'drug name' in tr.get_text().lower():
                        pass
                    else:
                        all_trs.append(tr)
            except:
                for tr in table.find('thead').findAll('tr'):
                    if 'drug name' in tr.get_text().lower():
                        pass
                    else:
                        all_trs.append(tr)
    count = 0
    for tr in all_trs:
        if 'Sponsor’s Comment:' in tr.get_text():
            pass
        else:
            all_tds = tr.findAll('td')
            drug_name = all_tds[0].find('p').find('span').get_text()
            try:
                is_drug = all_tds[0].find('p').findAll('span')[1].get_text()
            except:
                is_drug = all_tds[0].findAll('p')[1].find('span').get_text()

            if drug_name.isupper():
                drug_name = drug_name
            elif tr.find('td').find('ol'):
                ols = tr.find('td').findAll('ol')
                for ol in ols:
                    drug_name =drug_name + ' ' + ol.get_text()
            else:
                pass

            if is_drug.isupper():
                drug_name = drug_name + ' ' + is_drug
            drug_name1 = drug_name.strip().replace('Sponsor’S Comment:','')
            drug_name = drug_name.strip().title().replace('Sponsor’S Comment:','')
            
            drug_nam.append(drug_name)
#--------------------------------------------------------------------------------------------------------
            trade_name = ''
            trades = []
            check_p = all_tds[0].findAll('p')
            if len(check_p)>1:
                for ch_p in check_p:
                    for each_sp in ch_p.findAll('span'):
                        if '®' in each_sp.get_text():
                            trade_name =  trade_name + '\n' + each_sp.get_text()
                            trades.append(each_sp.get_text())
            else:
                trade_check = all_tds[0].find('p').findAll('span')
                for td_check in trade_check:
                    if '®' in td_check.get_text():
                        trade_name = trade_name + '\n' + td_check.get_text()
                        trades.append(each_sp.get_text())
            try:
                trade_name = trade_name.strip()
            except:
                trade_name = ' '
            trade_nam.append(trade_name)
#-------------------------------------------------------------------------------------------------------
            type_sub = ''
            for ch_p in check_p:
                for each_sp in ch_p.findAll('span'):
                    if '(Other' in each_sp.get_text():
                        type_submission = each_sp.get_text()
                        type_sub = type_submission
                    if 'submission' in each_sp.get_text().lower():
                        type_submission = each_sp.get_text()
                        type_sub = type_submission
                    if '(other' in each_sp.get_text():
                        type_submission = each_sp.get_text()
                        type_sub = type_submission
            type_submission = type_sub
            type_submissio.append(type_submission)
#------------------------------------------------------------------------------------------------------
            listin = ''
            for ch_p in check_p:
                for each_sp in ch_p.findAll('span'):
                    if 'listing' in each_sp.get_text().lower():
                        listing = each_sp.get_text()
                        listin = listing + listin
                    if 'listed' in each_sp.get_text().lower():
                        listing = each_sp.get_text()
                        listin = listing + listin
            # listing = listin
            listins.append(listin)
#-----------------------------------------------------------------------------------------------------
            sponso = ''
            for ch_p in check_p:
                for al_sp in ch_p.findAll('span'):
                    if 'Ltd' in al_sp.get_text():
                        sponsor = al_sp.get_text()
                        sponso = sponsor + sponso
                    if 'ltd' in al_sp.get_text():
                        sponsor = al_sp.get_text()
                        sponso = sponsor + sponso
                    if 'limit' in al_sp.get_text():
                        sponsor = al_sp.get_text()
                        sponso = sponsor + sponso
                    if 'Limit' in al_sp.get_text():
                        sponsor = al_sp.get_text()
                        sponso = sponsor + sponso
            # sponsor = sponso
            sponsos.append(sponso)
#-------------------------------------------------------------------------------------------------------
            form_text = ''
            texts = all_tds[0].findAll('p')
            for text in texts:
                for each_sp in text.findAll('span'):
                    form_text = form_text + ' ' + each_sp.get_text().replace('\xa0','')
                    form_text =form_text.replace('\xa0','')
            formek = [form_text]
            for trad in trades:
                form = formek[-1].replace(trad,'').replace(drug_name1,'').replace(sponsor,'').replace(listing,'').replace(type_submission,'').replace(u'\xa0','').strip()
                formek.append(form)
            forme.append(form)

            drug_use = all_tds[1].get_text()
            drug_us.append(drug_use)
            requested_listing = all_tds[2].get_text()
            requested_listin.append(requested_listing)
            if len(all_tds)<5:
                pbac_outcome = '-'
            else:
                pbac_outcome = all_tds[3].get_text()
            pbac_outcom.append(pbac_outcome)
            pbac_recom = all_tds[-1].get_text()
            pbac_reco.append(pbac_recom)
        count = count + 1
        # print(count)
        try:
            if 'Sponsor’s Comment:' in all_trs[count].findAll('td')[0].get_text():
                pass
            elif 'Sponsor’s Comment:' in tr.findAll('td')[0].get_text():
                try:
                    sponsor_comment = tr.findAll('td')[0].get_text().replace('Sponsor’s Comment:','')
                except:
                    sponsor_comment = '-'
                sponsor_commen.append(sponsor_comment)
                data = {
                    'Type': 'PBAC Web Outcome',
                    'Drug Name': drug_nam[-1],
                    'Form': forme[-1].strip(),
                    'Trade Name': trade_nam[-1].strip(), 
                    'Sponsor': sponsos[-1].strip(),
                    'Listing': listins[-1].strip(),
                    'Type Submission': type_submissio[-1].strip(),
                    'Drug Use': drug_us[-1].strip(),
                    'Requested Listing': requested_listin[-1].strip(),
                    'PBAC Outcome': pbac_outcom[-1].strip(),
                    'PBAC Recommendation': pbac_reco[-1].strip(),
                    'Sponsor’s comment': sponsor_commen[-1].strip()
                }
                lists.append(data)
            else:
                data = {
                    'Type': 'PBAC Web Outcome',
                    'Drug Name': drug_nam[-1],
                    'Form': forme[-1].strip(),
                    'Trade Name': trade_nam[-1].strip(), 
                    'Sponsor': sponsos[-1].strip(),
                    'Listing': listins[-1].strip(),
                    'Type Submission': type_submissio[-1].strip(),
                    'Drug Use': drug_us[-1].strip(),
                    'Requested Listing': requested_listin[-1].strip(),
                    'PBAC Outcome': pbac_outcom[-1].strip(),
                    'PBAC Recommendation': pbac_reco[-1].strip(),
                    'Sponsor’s comment': '-'
                }
                lists.append(data) 

            
        except:
            pass
    df = pd.DataFrame(lists).reset_index(drop=True)
    df.to_csv(f'{link_test}.csv',encoding = 'utf-8-sig',index=False)
    print(f'''Done...Look for a csv file : {link_test}
location: {os.getcwd()}''')


    
def mainmenu():
    print('1. Scrape Agenda doc')
    print('2. Positive Recommendation doc')
    print('3. First Time Decisions doc')
    print('4. Subsequest decisions doc')
    print('5. deferrals doc')
    print('6. other matters doc')
    print('7. PBAC Web Outcomes doc ')
    print('8. Exit')
    check = int(input('Enter Choice: '))
    if check == 1:
        scrape_agenda()
    elif check == 2:
        positive_recom()
    elif check == 3:
        first_time_decisions()
    elif check == 4:
        subsequent_decision()
    elif check == 5:
        scrape_deferral()
    elif check == 6:
        other_matters()
    elif check == 7:
        web_outcomes()
    elif check == 8:
        exit()
    else:
        print("Invalid Choice. Inter 1-8")
        mainmenu()

mainmenu()