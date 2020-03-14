import urllib.request
import json
from html.parser import HTMLParser
from bs4 import BeautifulSoup





serials = '07864D631,099880631,100247631,117809631,118187631,118898631,LVNJE7631,LVNJWJ631,LVNKDF631'
serials = serials.split(",")

for x in  serials:
    try:
        url = "https://w3-services1.w3-969.ibm.com/myw3/unified-profile/v1/docs/instances/master?userId=" + x  #By Serial Number
        # url = "https://w3-services1.w3-969.ibm.com/myw3/unified-profile/v1/docs/instances/masterByEmail?email=" + x    #By email
        
        
        contents = urllib.request.urlopen(url).read()
        # print(" testess")
        # soup = BeautifulSoup(contents,"html.parser")
        obj  =  json.loads(contents)
    except:
       print (x + "|Not Found")         
       continue


    try:
        mgr = obj["content"]["identity_info"]["functionalManager"]["preferredIdentity"]
    except:
        mgr = ""


    try:
        name = obj["content"]["identity_info"]["nameFull"]
    except:
        name = ""

    try:
        practice = obj["content"]["identity_info"]["practiceAlignment"]
    except:
        practice = ""

    try:
        email = obj["content"]["identity_info"]["mail"][0]
    except:
        email = ""


    try:
        country = obj["content"]["identity_info"]["co"]
    except:
        country = ""


    
    try:
        country = obj["content"]["identity_info"]["co"]
    except:
        country = ""


    try:
        role = obj["content"]["identity_info"]["role"]
    except:
        role = ""

    try:
        city = obj["content"]["identity_info"]["address"]["business"]["locality"]
    except:
        city = ""

    myList = [x, name, practice, email, city, role ]
    # myList = [x, mgr]
    csvLine = "|".join(myList)
    print (csvLine )

