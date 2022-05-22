import requests

url = "https://microsoft-translator-text.p.rapidapi.com/Detect"

querystring = {"api-version":"3.0"}

payload = """[\r
{\r
\"Text\": \"Ich würde wirklich gern Ihr Auto um den Block fahren ein paar Mal.\"\r
}\r
]"""
headers = {
'content-type': "application/json",
'x-rapidapi-key': "926576883emsh0929c234d91e5d8p19a217jsn1b07e87e770d",
'x-rapidapi-host': "microsoft-translator-text.p.rapidapi.com"
}

response = requests.request("POST", url, data=payload, headers=headers, params=querystring)

print(response.text)
