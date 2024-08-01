import requests

url = "https://ap.ceec.edu.tw/RegExam/ScoreSearch/Login?examtype=B"
 

ID = "21035415"
PID = "L125715002"

data = {
    "TestID": ID,
    "PID": PID,
    "Captcha": "19",
    "ExamType":	"B",
    "__RequestVerificationToken": "B89si7ae-mxqZZwkCydx-3tOg17OVliRqeTcsmxHuaWebmQtJe0y_FyJ83YyrrPxc8Uy9580sAnj0gzlUqIzy4e2d1TMRekL05KsCNEMwsw1"
}

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
           }

session = requests.Session()

initial_response = session.get(url)

# print(initial_response.text)
# captcha_solution = ""

# 19
# TS011ec7b5 = "01e5022d03f9d1a8e9f01bc48bc4162506d6f1dadcc590a4bf308c83fca1223eee7dd27f7c7024f0af380050c18fdf92066d9947e98eb84b9061b1ceb5d660c36efd71ad110019035338bc3a1f72c53729dbbb05a3f99a21cb1f46b54bb08700850e47b0cde55dcb6b58cf377223c516bef01c281d2735e721f62d616d8177b01d33d06fed93bb11c45545500109e1bd123c112a91a3ddb6a86bfc58aa9c8ae61ce9a503d4"

# session.cookies.set('TS011ec7b5', TS011ec7b5)


# response = session.post("https://ap.ceec.edu.tw/RegExam/ScoreSearch/SignIn", data=data, headers=headers)
# response = requests.post("https://ap.ceec.edu.tw/RegExam/ScoreSearch/SignIn")

# print(response.status_code)
print(response.text)