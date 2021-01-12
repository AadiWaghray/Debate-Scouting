from docx import Document
import re
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

main_doc_path = 'E:\Dropbox\Dropbox\Tournament Prep\ASU.docx'

main_doc = Document(main_doc_path)
styles = main_doc.styles
new_paragraph = main_doc.add_paragraph

list_of_entries = []
TBA = []
no_wiki = []

#Initialize the Web Driver
browser = webdriver.Chrome('Documents\Code\Python\chromedriver.exe')
wait = WebDriverWait(browser, 10)




#Store schools
schools = []
school_not_on_wiki= []

#Store teams


argument_list = []
recorded_schools = ['filler']
issue_words = ['High', 'School', 'The', 'HS', 'College Preparatory', 'BASIS', 'Of', 'Sr', 'Preparatory', 'Independent']

#Method to clean up school names
def RemoveBannedWords(toPrint,database):
    database_1 = sorted(list(database), key=len)
    pattern = re.compile(r"\b(" + "|".join(database_1) + ")\\W", re.I)
    return pattern.sub("", toPrint + ' ')[:-1]

#Navigate to tournament page
browser.get('https://www.tabroom.com/index/tourn/fields.mhtml?tourn_id=18419&event_id=156454')
browser.maximize_window()

#Retrieve number of entries
number_of_entries = browser.find_element_by_xpath('//*[@id="content"]/div[2]/span[2]/h5').text
number_of_entries = int(re.findall('\d+', number_of_entries)[0])



for counter in range(number_of_entries):
    
    #Clean team names (remove '&' along with whitespace)
    team_name = re.sub(r"\s+", "", browser.find_element_by_xpath(f'//*[@id="fieldsort"]/tbody/tr[{counter + 1}]/td[3]').text.replace('&', '-'))

    #Clean school names (take out words that will cause problem on wiki)
    print(counter)
    school_name = RemoveBannedWords(browser.find_element_by_xpath(f'//*[@id="fieldsort"]/tbody/tr[{counter + 1}]/td[1]').text, issue_words)

    #Removes TBA teams
    if  team_name == 'NamesTBA':
        TBA.append(school_name)

    else:
        #Hardcoding ADL exception to cleaning since it is an abbreviation
        if school_name == 'ADL':
            schools.append('Asian')

        else: 
            schools.append(school_name)
    list_of_entries.append(team_name)

list_of_entries.remove('Koh-Tsai')# This team has some problems with their wiki for some reason that makes indexing hard and breaks the program
schools.remove('Asian')#To make sure that both lists are still lined up





#Number of valid entries
number_of_valid_entries = len(schools)

aff_title = new_paragraph('***Affs***')
aff_title.style = styles['Heading 1']

for counter in range(number_of_valid_entries):
    #Navigate/ reset to wiki
    browser.get('https://hspolicy.debatecoaches.org/Main/')   

    #Checks if the school is on the wiki
    if len(browser.find_elements_by_xpath(f"//*[contains(text(), '{schools[counter]}')]")) > 0:
        #Waits untile element is clickable before clicking
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), '{schools[counter]}')]")))
        browser.find_element_by_xpath(f"//*[contains(text(), '{schools[counter]}')]").click()

        #Makes sure the name of the school is written only once on the document
        if not schools[counter] == recorded_schools[-1]:
            title = new_paragraph(schools[counter])
            title.style = styles['Heading 1']

            recorded_schools.append(schools[counter])
    else:
        school_not_on_wiki.append(schools[counter])

    #reverses entry's names
    split_entry = list_of_entries[counter].split('-')
    if  not split_entry == ['NamesTBA']:
        reverse_entry = f'{split_entry[1]}-{split_entry[0]}'
    else: 
        reverse_entry = 'NOTHING'

    #Current iteration entry
    current_entry = list_of_entries[counter]

    #Checks if entry is on wiki
    if len(browser.find_elements_by_xpath(f"//*[contains(text(), '{current_entry} Aff')]")) > 0 or len(browser.find_elements_by_xpath(f"//*[contains(text(), '{reverse_entry} Aff')]")) > 0:
        #Clicks on entry after it has loaded then write the name in the document
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), '{current_entry} Aff')]")))
            browser.find_element_by_xpath(f"//*[contains(text(), '{current_entry} Aff')]").click()
        except:
            wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), '{reverse_entry} Aff')]")))
            browser.find_element_by_xpath(f"//*[contains(text(), '{reverse_entry} Aff')]").click()
        finally:
            entries_paragraph = new_paragraph(current_entry)
            entries_paragraph.style = styles['Heading 2']
            new_paragraph(browser.current_url)

            round_reports_title = new_paragraph('Round Reports')
            round_reports_title.style = styles['Heading 3']
            
            #The following for loops iterate through a table until there are no more elements or reach the arbitrary range I set

            #Iterate through round report table
            for i in range(30):
                
                #Checks if there is a round report
                if len(browser.find_elements_by_xpath(f'//*[@id="tblReports"]/tbody/tr[{i + 2}]')) > 0:
                    round_report = browser.find_element_by_xpath(f'//*[@id="tblReports"]/tbody/tr[{i + 2}]').find_element(By.NAME, 'report').get_attribute('innerHTML').replace('<p>', "").replace('<br>', " ").replace('</p>', "")
                    round = browser.find_element_by_xpath(f'//*[@id="tblRounds"]/tbody/tr[{i +2}]').text

                    round_entry = new_paragraph(round)
                    round_entry.style = styles['Heading 4']

                    round_report_entry = new_paragraph(round_report)
                else: 
                    break

            argument_title = new_paragraph('Arguments on wiki')
            argument_title.style = styles['Heading 2']

            #Iterates thorugh argument table
            for i in range(20):
                #Checks if there is an argument
                if len(browser.find_elements_by_id(f'title{i}')) > 0:
                    argument_name = new_paragraph(browser.find_element_by_id(f'title{i}').text)
                    argument_name.style = styles['Heading 3']
                    
                    argument = new_paragraph(browser.find_element_by_id(f'entry{i}').get_attribute('textContent'))
                else:
                    break

            file_title = new_paragraph('Files')
            file_title.style = styles['Heading 3']

            #Iterates through files
            for i in range(20):
                if len(browser.find_elements_by_xpath(f'//*[@id="tblOpenSource"]/tbody/tr[{i + 2}]')) > 0:
                    url = browser.find_element_by_xpath(f'//*[@id="tblOpenSource"]/tbody/tr[{i + 2}]/td[1]/div/p/span/a').get_attribute('href')

                    r = requests.get(url)

                    source_doc_path = f'E:\Dropbox\Dropbox\Tournament Prep\File dump\Aff {schools[counter]}_{list_of_entries[counter]}_{i}.docx'
                    with open(source_doc_path, 'wb') as f:
                        f.write(r.content)
                    
                    round_reference = new_paragraph(browser.find_element_by_xpath(f'//*[@id="tblOpenSource"]/tbody/tr[{i + 2}]').text)
                    round_reference.style = styles['Heading 4'] 
                    
                    new_paragraph(f'file://{source_doc_path}')
                else:
                    break

    else:
        no_wiki.append(current_entry)
    main_doc.save(main_doc_path)





#Need to finish comments and cleaning from this point on
neg_title = new_paragraph('***Negs***')
neg_title.style = styles['Heading 1']

for counter in range(number_of_valid_entries):
    browser.get('https://hspolicy.debatecoaches.org/Main/')  

    if len(browser.find_elements_by_xpath(f"//*[contains(text(), '{schools[counter]}')]")) > 0:
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), '{schools[counter]}')]")))
        browser.find_element_by_xpath(f"//*[contains(text(), '{schools[counter]}')]").click()

        if schools[counter] == recorded_schools[-1]:
            print('School title is already in document')
        else:
            title = new_paragraph(schools[counter])
            title.style = styles['Heading 1']

            recorded_schools.append(schools[counter])
    else:
        school_not_on_wiki.append(schools[counter])

    split_entry = list_of_entries[counter].split('-')
    print(split_entry)
    if  not split_entry == ['NamesTBA']:
        reverse_entry = f'{split_entry[1]}-{split_entry[0]}'
    else: 
        reverse_entry = 'NOTHING'

    if len(browser.find_elements_by_xpath(f"//*[contains(text(), '{list_of_entries[counter]} Neg')]")) > 0 or len(browser.find_elements_by_xpath(f"//*[contains(text(), '{reverse_entry} Neg')]")) > 0:
        try:
            browser.find_element_by_xpath(f"//*[contains(text(), '{list_of_entries[counter]} Neg')]").click()
        except:
            browser.find_element_by_xpath(f"//*[contains(text(), '{reverse_entry} Neg')]").click()
        finally:
            new_paragraph(browser.current_url)
            entries_paragraph = new_paragraph(list_of_entries[counter])
            entries_paragraph.style = styles['Heading 2']

            round_reports_title = new_paragraph('Round Reports')
            round_reports_title.stlye = styles['Heading 3']
            for i in range(30):
                if len(browser.find_elements_by_xpath(f'//*[@id="tblReports"]/tbody/tr[{i + 2}]')) > 0:
                    round_report = browser.find_element_by_xpath(f'//*[@id="tblReports"]/tbody/tr[{i + 2}]').find_element(By.NAME, 'report').get_attribute('innerHTML').replace('<p>', "").replace('<br>', " ").replace('</p>', "")
                    round = browser.find_element_by_xpath(f'//*[@id="tblRounds"]/tbody/tr[{i +2}]').text

                    round_entry = new_paragraph(round)
                    round_entry.style = styles['Heading 4']

                    round_report_entry = new_paragraph(round_report)
                else: 
                    break

            argument_title = new_paragraph('Arguments on wiki')
            argument_title.style = styles['Heading 2']

            for i in range(20):
                if len(browser.find_elements_by_id(f'title{i}')) > 0:
                    

                    argument_name = new_paragraph(browser.find_element_by_id(f'title{i}').text)
                    argument_name.style = styles['Heading 3']
                    
                    argument = new_paragraph(browser.find_element_by_id(f'entry{i}').get_attribute('textContent'))
                else:
                    break
            file_title = new_paragraph('Files')
            file_title.style = styles['Heading 3']
            for i in range(20):
                if len(browser.find_elements_by_xpath(f'//*[@id="tblOpenSource"]/tbody/tr[{i + 2}]')) > 0:
                    url = browser.find_element_by_xpath(f'//*[@id="tblOpenSource"]/tbody/tr[{i + 2}]/td[1]/div/p/span/a').get_attribute('href')

                    r = requests.get(url)

                    source_doc_path = f'E:\Dropbox\Dropbox\Tournament Prep\File dump\Aff {schools[counter]}_{list_of_entries[counter]}_{i}.docx'
                    with open(source_doc_path, 'wb') as f:
                        f.write(r.content)
                    
                    round_reference = new_paragraph(browser.find_element_by_xpath(f'//*[@id="tblOpenSource"]/tbody/tr[{i + 2}]').text)
                    round_reference.style = styles['Heading 4'] 
                    
                    new_paragraph(f'file://{source_doc_path}')
                else:
                    break

            print('done with entry')

    main_doc.save(main_doc_path)



a = new_paragraph('These teams do not have a wiki page')
a.style = styles['Heading 1']
for i in range(len(no_wiki)):
    a = new_paragraph(no_wiki[i])
    a.style = styles['Heading 4']

a = new_paragraph('This is how many teams do not have an entry determined: ' + str(len(TBA)) + '. They are....')
a.style = styles['Heading 1']
for i in range(len(TBA)):
    a = new_paragraph(TBA[i])
    a.style = styles['Heading 4']

a = new_paragraph('This is how many schools are not on the wiki: ' + str(len(school_not_on_wiki)) + '. They are....')
a.style = styles['Heading 1']
for i in range(len(school_not_on_wiki)):
    a = new_paragraph(school_not_on_wiki[i])
    a.style = styles['Heading 4']





main_doc.save('E:\Dropbox\Dropbox\Tournament Prep\ASU.docx')

browser.close()

#TO-DO
#-https://www.selenium.dev/documentation/en
#-https://python-docx.readthedocs.io/en/latest/index.html#api-documentation 
#-Write the team's arguments into a verbatim document WITH formatting
#Create methods do clean up the code and comment it out
#Make a lot of the file path stuff automated and exportable for application
#Deal with joh team weirdness