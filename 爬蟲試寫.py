import requests as req
from openpyxl import Workbook 


wb = Workbook()
ws = wb.active

title =["課名", "作者", "價格", "預購價", "販售數"]
ws.append(title)

header = {
    "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Safari/537.3"
}

for index in range(28):
    url = "https://api.hahow.in/api/courses?limit=24&page=0"
    url = url + str(index)
    print(url)
    r = req.get(url, headers=header)
    print(r)

    root_json = r.json()

    for data in root_json["data"]:
        course = []
        course.append(data["title"])
        course.append(data["owner"]["name"])
        course.append(data["price"])
        course.append(data["preOrderedPrice"])
        course.append(data["numSoldTickets"])

        ws.append(course)

wb.save("data.xlsx")


 