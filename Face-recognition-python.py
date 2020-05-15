# import modules
import http.client, urllib.request, urllib.parse, urllib.error, base64, json
from PIL import Image, ImageDraw
import requests
from PIL import Image
from datetime import datetime
import xlwt
import glob
import cv2
import sys
import os
import json
import numpy as np

#-------------------------------

KEY = 'b6ffdfe0c6564f5fb4bedc9600006588'
headers = {
    'Content-Type': 'application/json',
    'Ocp-Apim-Subscription-Key': KEY,
}
#------------------------------


###PersonGroup — Create

personGroupId="employees"
##
##body = dict()
##body["name"] = "Employees"
##body["userData"] = "All Employees"
##body = str(body)
##
###Request URL 
##FaceApiCreateLargePersonGroup = 'https://recognition-face.cognitiveservices.azure.com/face/v1.0/persongroups/'+personGroupId 
##
##try:
##    # REST Call 
##    response = requests.put(FaceApiCreateLargePersonGroup, data=body, headers=headers) 
##    print("RESPONSE:" + str(response.status_code))
##
##except Exception as e:
##    print(e)

#-------------------------------


### PersonGroup Person — Create

#Request Body
body = dict()
body["name"] = "Malak"
body["userData"] = "IT Department"
body = str(body)
  

#Request URL 
FaceApiCreatePerson = 'https://recognition-face.cognitiveservices.azure.com/face/v1.0/persongroups/'+personGroupId+'/persons' 

try:
    # REST Call 
    response = requests.post(FaceApiCreatePerson, data=body, headers=headers)
  
    responseJson = response.json()
    personId = responseJson["personId"]
    #print("PERSONID: "+str(personId)) 
    
except Exception as e:
    print(e)

#--------------------------------

     
###PersonGroup Person — Add Face
 

data=[]
directory = r'C:/Users\Admin\AppData\Local\Programs\Python\Python37\Azure project\malak'
for filename in os.listdir(directory):
    if filename.endswith(".jpg") :
        data.append(os.path.join(directory, filename))
    else:
        continue

#print(data)

    

    
####Request URL 
FaceApiCreatePerson = 'https://recognition-face.cognitiveservices.azure.com/face/v1.0/persongroups/'+personGroupId+'/persons/'+personId+'/persistedFaces' 

for image in data:
    #w = open(image, 'r+b')
    body = dict()
    body["url"] = image
    body = str(body)

    try:
        # REST Call 
        response = requests.post(FaceApiCreatePerson, data=body, headers=headers)
        responseJson = response.json()
        PersistedFaceId = responseJson["persistedFaceId"]
        
        #print("PERSISTED FACE ID: "+ str(PersistedFaceId))
        
    
    except Exception as e:
        print('')

#------------------------------------  

### PersonGroup  — Train
        
##Request Body
body = dict()

##Request URL 
FaceApiTrain = 'https://recognition-face.cognitiveservices.azure.com/face/v1.0/persongroups/'+personGroupId+'/train'

try:
    # REST Call 
    response = requests.post(FaceApiTrain, data=body, headers=headers)

    #print("RESPONSE:" + str(response.status_code))

except Exception as e:
    print(e)


#---------------------------

### setup the camera

currentframe = 0
cam = cv2.VideoCapture(0)  
while(True): 
     # reading from frame 
      ret,frame = cam.read() 
      if ret: 
        # if video is still left continue creating images 
         name = './data/frame' + str(currentframe) + '.jpg'
        # print ('Creating...' + name)
         cv2.imshow("preview", frame)
         
        # writing the extracted images 
         cv2.imwrite(name, frame)  
   
         currentframe += 1
      else: 
         break
      k = cv2.waitKey(30) & 0xff 
    
      if k == 27: # press 'ESC' to quit
           break
      elif currentframe>=30:
         break
# Release all space and windows once done 
cam.release() 
cv2.destroyAllWindows() 

#----------------------------

### Calculate date and time

String1=[]
String2=[]
String3=[]

now=datetime.now()
date=now.strftime('%d/%m/%Y')
time=now.strftime('%H:%M:%S')

if time > str(19) :
    string1="Arrived Late"
    String1.append(string1)
elif time < str(19) :
    string2='Arrived early'
    String2.append(string2)
else:
    string3='On time'
    String3.append(string3)
    
   
      

#----------------------------
    
### Face — Detect 

faceIdList=[]
Face=[]

def detect(img_url):
    headers = {'Content-Type': 'application/octet-stream', 'Ocp-Apim-Subscription-Key': KEY}
    body = open(img_url,'rb')
    params = urllib.parse.urlencode({'returnFaceId': 'true','returnFaceAttributes': 'age,gender,glasses,noise'})

    conn = http.client.HTTPSConnection('recognition-face.cognitiveservices.azure.com')
    conn.request("POST", '/face/v1.0/detect?%s' % params, body, headers)
    response = conn.getresponse()
    photo_data = json.loads(response.read())
    #print(photo_data)
    
    if not photo_data: # if post response is empty (no face found)
        print('No face Detected')
    else: # if face is found
        for face in photo_data: # for the faces identified in each photo
            
            faceIdList.append(str(face['faceId'])) # get faceId for use in identify
            Face.append(str(face['faceAttributes']))
            
            
detect(name)



#-----------------------------

### Face — Identify


#Request Body

Name=[]
Confidence=[]
Date=[]
Time=[]
personID=[]

body = dict()
body["personGroupId"] = personGroupId
body["faceIds"] = faceIdList
body = str(body)

workbook=xlwt.Workbook(encoding='utf-8')
sheet1=workbook.add_sheet('Employees Attendance')
sheet1.write(0,0,"ID")
sheet1.write(0,1,"Name")
sheet1.write(0,2,"Date")
sheet1.write(0,3,"Time")
sheet1.write(0,4,"Confidace")
sheet1.write(0,5,"Status")

sheet2=workbook.add_sheet('New persons')
sheet2.write(0,0,"Attributes")
sheet2.write(0,1,"Date")
sheet2.write(0,2,"Time")



# Request URL 
FaceApiIdentify = 'https://recognition-face.cognitiveservices.azure.com/face/v1.0/identify' 
FaceApiGetPerson = 'https://recognition-face.cognitiveservices.azure.com/face/v1.0/persongroups/'+personGroupId+'/persons/'+personId

try:
    # REST Call 
    response = requests.post(FaceApiIdentify, data=body, headers=headers) 
    responseJson = response.json()
    personId = responseJson[0]["candidates"][0]["personId"]
    confidence = responseJson[0]["candidates"][0]["confidence"]
    if (confidence > .7):
        response = requests.get(FaceApiGetPerson, headers=headers)
        responseJson = response.json()
        print("Successfully identified")
        Name.append(responseJson["name"])
        personID.append(personId)
        Confidence.append(confidence)
        Date.append(date)
        Time.append(time)
     

        
        for item1 in range(len(Name)):
             sheet1.write(item1+1,0,str(personID[item1]))
             sheet1.write(item1+1,1,str(Name[item1]))
             sheet1.write(item1+1,2, str(Date[item1]))
             sheet1.write(item1+1,3, str(Time[item1]))
             sheet1.write(item1+1,4, str(Confidence[item1]))
             sheet1.write(item1+1,5, str(String1[item1]))
             sheet1.write(item1+1,5, str(String2[item1]))
             sheet1.write(item1+1,5, str(String3[item1]))
             
        print('Done')
             
        
        
    else:
           print('None')
        
except Exception as e:
    print("Could not identify new person")
    Date.append(date)
    Time.append(time)
    for item1 in range(len(Face)):
        sheet2.write(item1+1,0, str(Face[item1]))
        sheet2.write(item1+1,1, str(Date[item1]))
        sheet2.write(item1+1,2, str(Time[item1]))
    print('Done')
    

    
workbook.save('Attendace.xls') 





