import requests
import lxml
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import pandas as pd
from requests.exceptions import ChunkedEncodingError
from requests.exceptions import ConnectionError
from requests.exceptions import ReadTimeout
import time
import re
import datetime
from bs4 import SoupStrainer

#File Path
input_File_Path = "Linkedin.xlsx"
output_File_Path = "output.xlsx"

#define the range from 0
startIndex = 0
endIndex = 10

# Below Strings are the base for URL creation
string1 = "https://www.bing.com/search?q="
string2 = "+site:linkedin.com/in"
string3 = "&form=QBLH"

# below function to store the data into excel file
file_path = output_File_Path
wb = load_workbook(file_path)
sheet = wb.active

# below function is to read the excel data
pd.read_excel(input_File_Path )
var = pd.read_excel(input_File_Path , index_col=0, usecols='B', sheet_name=0)
company_names = var.index

# below function is to read the excel data
pd.read_excel(input_File_Path )
var2 = pd.read_excel(input_File_Path , index_col=0, usecols='C', sheet_name=0)
titles = var2.index

# below function is to read the excel data
pd.read_excel(input_File_Path )
var3 = pd.read_excel(input_File_Path , index_col=0, usecols='A', sheet_name=0)
code = var3.index

def remove_non_ascii(input_str):
    """
    Remove non-ASCII characters from a string.
    """
    return re.sub(r'[^\x00-\x7F]+', '', input_str)

len_of_data = len(company_names) # Will calculate the length
serial = 1 # Will throw serial according to the loop
s = requests.Session()
start = datetime.datetime.now()
for i in range(startIndex,endIndex):
    # Throw output as Company name and serial no 
    time.sleep(1)
    try:
        codes = code[i]
        input_string = company_names[i]
        company_name = remove_non_ascii(input_string)
        title = titles[i]
        for r in ((",",""), ("/",""), (";",""), ("&",""), (" ","+")):
            company_name = company_name.replace(*r)
            title = title.replace(*r)
        url = string1 + company_name + title + string2 + string3 # URL is generated
        print(str(i) + " " + company_name.replace("+"," "))
        print(f"Batch No :{codes}")
        # below function is to execute the url with the help og beautiful soup
        headers={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36"}
        r = s.get(url,headers=headers,timeout=2)
        only_tags = SoupStrainer("ol",id="b_results")
        soup = BeautifulSoup(r.content, 'lxml', parse_only=only_tags)

        if r.status_code == 200:
            
            descriptions = soup.find_all("p",class_="b_algoSlug") # Scrape all the paragraph element
            container = soup.findAll("div",class_="b_vlist2col") # Scrape all the container element
            head_container = soup.find_all("li",class_="b_algo") # Scrape all the list element

            links = soup.findAll("div",class_="b_attribution")
            len_links = len(links) # calculate the length
            if len_links == 0:
                # loop(i,serial)
                data = {"Serial":"","Formula":"","Formula2":"","code":"","Input":"","Input2":"","Name":"","Link":"","Header":"","Title":"","Work":"","Connection":"","Followers":"","Location":"","Description":""}
                list_a = []
                data['Serial'] = ""
                data['Formula'] = ""
                data['Formula2'] = ""
                data['code'] = codes
                data['Input'] = company_name.replace("+"," ")
                data['Input2'] = title.replace("+"," ")
                data['Name'] = url
                data['Link'] = "There is no result"
                data['Header'] = ""
                data['Title'] = ""
                data['Work'] = ""
                data['Connection'] = ""
                data['Followers'] = ""
                data['Location'] = ""
                data['Description'] = ""
                data['Match'] = ""
                serial += 1

                pase1 = list(data.values())
                list_a.append(pase1)

                for item in list_a: # append the extracted data into sheet
                    sheet.append(item)
                
                wb.save(filename=output_File_Path) # save the file
            for count in range(len_links):
                try:
                    data = {"Serial":" ","Formula":" ","Formula2":" ","code":" ","Input":" ","Input2":" ","Name":" ","Match":"","Link":" ","Header":"","Title":" ","Work":" ","Connection":" ","Followers":" ","Location":" ","Description":" "}
                    list_a = []
                    data['Serial'] = ""
                    data['Formula'] = ""
                    data['Formula2'] = ""
                    data['code'] = codes
                    data['Input'] = company_name.replace("+"," ").replace('"','').upper()
                    data['Input2'] = title.replace("+"," ").replace('"','').upper()
                    each_link = links[count]
                    link = each_link.text
                    http,profile = link.rsplit('/',1)
                    try:
                        name,id = profile.rsplit('-',1)
                        data['Name'] = name.replace("-"," ")
                        data['Link'] = link
                    except:
                        data['Link'] = link
                        data['Name'] = profile
                    data['Description'] = descriptions[count].text.replace('Web','')
                    each_container = container[count]
                    li_container = each_container.findAll("li")
                    for li in li_container:
                        li_content = li.text 
                        if li_content.startswith("Title:"):
                            data['Title'] = li_content.replace("Title:","")
                        elif li_content.startswith("Connections:"):
                            data['Connection'] = li_content.replace("Connections:","") 
                        elif li_content.startswith("Location:"):
                            if "followers" in li_content:
                                data['Followers'] = li_content.replace("Location:","")  
                            else:
                                data['Location'] = li_content.replace("Location:","").replace('abonnés','followers').replace('mil seguidores','k followers').replace('seguidores','followers') 
                    head_content = head_container[count]

                    content = head_content.text.lower() # Extracted Text from the head container
                    replace_quote_plus = company_name.replace("+"," ").replace('"','')
                    keyword_string = replace_quote_plus.lower()
                    keyword_sub_string = keyword_string.split()
                    keyword_list = []
                    for j in range(len(keyword_sub_string)):
                        keyword_list.append(" ".join(keyword_sub_string[:j+1]))
                    is_matched = any(keyword in content for keyword in keyword_list[1:])
                    if is_matched:
                        data['Match'] = "Matched"
                    else:
                        data['Match'] = " "

                    header = head_content.h2.text 
                    data['Header'] = header
                    try:
                        name,work = header.split("-",1)
                        data['Work'] = work.replace("- LinkedIn","")
                    except:
                        data['Work'] = header.replace("- LinkedIn","")
                    
                    count += 1 
                    serial += 1 

                    pase1 = list(data.values())
                    list_a.append(pase1)

                    for item in list_a:
                        sheet.append(item)
                    
                    wb.save(filename=output_File_Path)      
                except:
                    data = {"Serial":"","Formula":"","Formula2":"","code":"","Input":"","Input2":"","Name":"","Link":"","Header":"","Title":"","Work":"","Connection":"","Followers":"","Location":"","Description":""}
                    list_a = []
                    data['Serial'] = ""
                    data['Formula'] = ""
                    data['Formula2'] = ""
                    data['code'] = codes
                    data['Input'] = company_name.replace("+"," ").replace('"','').upper()
                    data['Input2'] = title.replace("+"," ").replace('"','').upper()
                    data['Name'] = "An Error occured while scraping"
                    data['Link'] = ""
                    data['Header'] = ""
                    data['Title'] = ""
                    data['Work'] = ""
                    data['Connection'] = ""
                    data['Followers'] = ""
                    data['Location'] = ""
                    data['Description'] = ""
                    data['Match'] = ""
                    serial += 1

                    pase1 = list(data.values())
                    list_a.append(pase1)

                    for item in list_a:
                        sheet.append(item)
                    
                    wb.save(filename=output_File_Path)
        else:
            data = {"Serial":"","Formula":"","Formula2":"","code":"","Input":"","Input2":"","Name":"","Link":"","Header":"","Title":"","Work":"","Connection":"","Followers":"","Location":"","Description":""}
            list_a = []
            data['Serial'] = ""
            data['Formula'] = ""
            data['Formula2'] = ""
            data['code'] = codes
            data['Input'] = company_name.replace("+"," ").replace('"','').upper()
            data['Input2'] = title.replace("+"," ").replace('"','').upper()
            data['Name'] = "Response Code 4xx"
            data['Link'] = ""
            data['Header'] = ""
            data['Title'] = ""
            data['Work'] = ""
            data['Connection'] = ""
            data['Followers'] = ""
            data['Location'] = ""
            data['Description'] = ""
            data['Match'] = ""
            serial += 1

            pase1 = list(data.values())
            list_a.append(pase1)

            for item in list_a:
                sheet.append(item)
            
            wb.save(filename=output_File_Path)
    except ChunkedEncodingError:
        data = {"Serial":"","Formula":"","Formula2":"","code":"","Input":"","Input2":"","Name":"","Link":"","Header":"","Title":"","Work":"","Connection":"","Followers":"","Location":"","Description":""}
        list_a = []
        data['Serial'] = ""
        data['Formula'] = ""
        data['Formula2'] = ""
        data['code'] = codes
        data['Input'] = company_name.replace("+"," ").replace('"','').upper()
        data['Input2'] = title.replace("+"," ").replace('"','').upper()
        data['Name'] = "Raised ChunkedEncoding Error"
        data['Link'] = ""
        data['Header'] = ""
        data['Title'] = ""
        data['Work'] = ""
        data['Connection'] = ""
        data['Followers'] = ""
        data['Location'] = ""
        data['Description'] = ""
        data['Match'] = ""
        serial += 1

        pase1 = list(data.values())
        list_a.append(pase1)

        for item in list_a: # append the extracted data into sheet
            sheet.append(item)
        
        wb.save(filename=output_File_Path) # save the file
    except ConnectionError:
        print("Connection Error")
    except ReadTimeout:
        print("Request timed out")
# will throw completed after completion of work
finish = datetime.datetime.now() - start 
print("*****************Completed************************")
print(finish)