import docx
import requests

document = docx.Document('file.docx')

url = "https://microsoft-translator-text.p.rapidapi.com/translate"

querystring = {"to":"te","api-version":"3.0","profanityAction":"NoAction","textType":"plain"}


headers = {
    'content-type': "application/json",
    'x-rapidapi-host': "microsoft-translator-text.p.rapidapi.com",
    'x-rapidapi-key': "926576883emsh0929c234d91e5d8p19a217jsn1b07e87e770d"
    }


table = document.tables[0]
line = table.rows[4].cells[1].paragraphs[0].text
payload = """[\r
    {\r
        \"Text\": \" """ + line + """\"\r
    }\r
]"""
response = requests.request("POST", url, data=payload.encode('utf-8'), headers=headers, params=querystring)
table.rows[0].cells[2].add_paragraph(response.text)
print(response.text)

# for r in range(1, len(table.rows)):
#     line = table.rows[r].cells[1].paragraphs[0]
#     payload = {
#             "Text": line
#         }
#     response = requests.request("POST", url, data=payload, headers=headers, params=querystring)
#     table.rows[r].cells[2].add_paragraph(response)
document.save('file.docx')

# line = 'HI'
# payload = """[\r
#     {\r
#     \"Text\": \" """+ line + """\"\r
#     }\r
# ]"""
# headers = {
#         'content-type': "application/json",
#         'x-rapidapi-key': "926576883emsh0929c234d91e5d8p19a217jsn1b07e87e770d",
#         'x-rapidapi-host': "microsoft-translator-text.p.rapidapi.com"
#         }
# response = requests.request("POST", url, data=payload.encode('utf-32'), headers=headers, params=querystring)       
# obj = (json.loads(response.text))
# text = obj[2]['translations'][0]['text']
# print(text)
# # table.cell(i,2).paragraphs[0].text = text


# /////////////////////////////
# if line == "":
#     i += 1
# runs = table.rows[i].cells[1].paragraphs[0].runs
# wrun = runs[len(runs) - 1]
# if wrun.font.color.rgb == None and wrun.font.highlight_color != WD_COLOR_INDEX.GRAY_25:
    
    
    
   
# ///////////////////////////// 
# payload = """[\r
# {\r
# \"Text\": \" """+ line + """\"\r
# }\r
# payload = [{"Text": "hi"}]
# headers = {
# 	"content-type": "application/json",
# 	"X-RapidAPI-Host": "microsoft-translator-text.p.rapidapi.com",
# 	"X-RapidAPI-Key": "926576883emsh0929c234d91e5d8p19a217jsn1b07e87e770d"
# }

# response = requests.post(url, json=payload, headers=headers, params=querystring)
# r = response.json()

# filee = input("type file name exact:")
# f = open(filee, "rb",  encoding ="utf8")

