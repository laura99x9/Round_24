
from bs4 import BeautifulSoup
import requests

url = 'https://org.uib.no/medborgernotat/kodeb%c3%b8ker/Codebook%20-%20NCP%20round%2024%20-%20v-100%20-%20en.html'
response = requests.get(url)
soup = BeautifulSoup(response.content, 'html.parser')
    
headline = soup.find('h4')
list_con = soup.find_all('h4') 
second = soup.find_all('tr')
old = str(second)
you = old.splitlines()


law = soup.prettify
ans = str(law).splitlines()
svar = []
for i in range(1550, len(ans)):
    if "<" not in ans[i]:
        if len(ans[i]) != 0:
            svar.append(ans[i])
           
test = []
for i in svar:
    if i[0] == 'r':
        test.append(i)

test_2 = []
other = ['Technical description:', 'Technical attributes:']
for i in range(len(svar)):
    if "question" in svar[i]:
        test_2.append(svar[i])
    if "question" in svar[i]:
        test_2.append(svar[i+1])
    for j in range(len(test)):
        if svar[i] == test[j]:
            test_2.append(svar[i])

#vis(test_2)

soupreme = []       
for i in test_2:
    if i[0] == 'r':
        soupreme.append(i)

#vis(soupreme)

love = []
for i in range(len(test_2)):
    for j in range(len(soupreme)):
        if test_2[i] == soupreme[j]:
            love.append(i)
#vis(love)

coco = []
for i in range(len(love)-1):
    for j in range(love[i], love[i+1]):
        coco.append(j)
#vis(coco)
test_3 = []
for i in coco:
    test_3.append(test_2[i])

test_4 = []
for i in range(len(test_3)):
    if test_3[i] != test_3[i-1] or test_3[i+1]:
        test_4.append(test_3[i])
        
#vis(test_4)

test_5_fake = []
#list_r = test
list_q = ['Pre-question text:','Literal question:', 'Post-question:']

for i in range(len(test_4)-1):
    if test_4[i] in test:
        test_5_fake.append(test_4[i])
    if test_4[i] in list_q:
        test_5_fake.append(test_4[i])
        test_5_fake.append(test_4[i+1])
    
test_5 = []

for i in range(len(test_5_fake)-1):
    if test_5_fake[i] != test_5_fake[i+1]:
        test_5.append(test_5_fake[i])

#vis(test_5)                                   
temp = [] 
for i in range(len(test_5)-1):
    if test_5[i+1] in test:
        temp.extend([test_5[i]])
    if test_5[i+1] not in test:
        temp.extend(["•" + test_5[i]])
        
if test_5[-1] not in test: 
    temp.extend(["•" + test_5[-1]])


amo = []#selve Spørsmålene
for i in range(len(temp)-1):
    if test_5[i] in test and test_5[i-1] not in test: amo.append(test_5[i-1])

#vis(amo)

out = []#r som har spørsmål + selve spørsmål

for i in range(len(temp)-1):
    if temp[i] in amo:
        out.append("•" + temp[i])
    else:
        out.append(temp[i])
#vis(out)
        
und = []#Alle r som har spørsmål
for i in out:
    if i[0:2] == "•r": und.append(i)


#vis(und)

result = []

i = 0
while i < len(out):
    if out[i] in und:
        line = out[i]
        i += 1
        while i < len(out) and (out[i].startswith("•") and out[i] not in und):
            line += " " + out[i]
            i += 1
        result.append(line)
    else:
        result.append(out[i])
        i += 1

#vis(result)

from openpyxl import Workbook
import os

# Create a new workbook
workbook = Workbook()
sheet = workbook.active

# Define the headers
headers = ['Items', 'Pre-Question', 'Literal question', 'Post-question']

# Write the headers to the first row of the sheet
sheet.append(headers)

# Iterate over the list
for item in result:
    # Check if the item starts with a bullet point
    if item.startswith('•'):
        # Extract the item name
        item_name = item.split('•')[1].strip()

        # Extract the question types and their corresponding headers
        question_types = {}
        question_parts = item.split('•')
        for i in range(1, len(question_parts)):
            question_part = question_parts[i].strip()
            if question_part.startswith('Pre-question text:'):
                question_types['Pre-question'] = question_parts[i+1].strip()
            elif question_part.startswith('Literal question:'):
                question_types['Literal question'] = question_parts[i+1].strip()
            elif question_part.startswith('Post-question:'):
                question_types['Post-question'] = question_parts[i+1].strip()

        # Write the item name and question types to a new row
        row_data = [item_name]
        row_data.append(question_types.get('Pre-question', ''))
        row_data.append(question_types.get('Literal question', ''))
        row_data.append(question_types.get('Post-question', ''))
        sheet.append(row_data)
    else:
        # Exclude bullet points and write non-bullet point items to a new row
        item_name = item.strip()
        sheet.append([item_name, '', '', ''])

# Save the workbook, replacing the existing file if it exists
file_name = 'round_24.xlsx'
if os.path.exists(file_name):
    os.remove(file_name)  # Remove existing file
workbook.save(file_name)

