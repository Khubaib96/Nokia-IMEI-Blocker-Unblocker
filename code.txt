import pyautogui
from appJar import gui
from win32com.client import Dispatch
import clipboard



#########################################




##################################################

screenWidth, screenHeight = pyautogui.size()
currentMouseX, currentMouseY = pyautogui.position()
wsh = Dispatch("WScript.Shell")

#############################################################################


def Block(btn):
    output = []
    no_of_imei = 0
    IMEI = clipboard.paste()
    IMEI = str(IMEI)
    IMEI = IMEI.replace("\r\n", ",")
    IMEI = IMEI.split(",")
    del IMEI[-1]


    header = "<spml:batchRequest" + "\n"  \
              "execution"+"="+'"synchronous"' + "\n"  \
              "processing"+"="+'"sequential"' + "\n"  \
              "onError"+"="+'"resume"' + "\n"  \
              "xmlns:spml"+"="+'"urn:siemens:names:prov:gw:SPML:2:0"' + "\n"  \
              "xmlns:nsr"+"="+'"urn:siemens:names:prov:gw:EIR_NSR:1:0"' + "\n"  \
              "xmlns:xsi"+"="+'"http://www.w3.org/2001/XMLSchema-instance">' + "\n"  \
              "<version>EIR_NSR_v30</version>" + "\n"

    initial_data = '<request xsi:type="spml:AddRequest">' + "\n"  \
                   '<version>EIR_NSR_v30</version>'+ "\n"  \
                   '<object xsi:type="nsr:Device">' + "\n"  \
                   "<identifier>"

    Final_data = '</identifier>' + "\n"  \
                 '<colour>' + "\n"  \
                 '<colour>black</colour>' + "\n"  \
                 '<reason>1</reason>' + "\n"  \
                 '<organization>ORG_1</organization>' + "\n"  \
                 '<deviceManufacturer>Nokia</deviceManufacturer>' + "\n"  \
                 '<deviceName>mobile</deviceName>' + "\n"  \
                 '<processedDate>2017-12-12</processedDate>' + "\n"  \
                 '<processedTime>12:01</processedTime>' + "\n"  \
                 '<duplicates>unique</duplicates>' + "\n"  \
                 '</colour>' + "\n"  \
                 '</object>' + "\n"  \
                 '</request>' + "\n"

    footer = "</spml:batchRequest>"

    output = header

    for imei in IMEI:
        output= output + initial_data + str(IMEI[no_of_imei]) + Final_data
        no_of_imei = no_of_imei + 1

    output = output + footer

    myfile = open("Blocking.spml", "w+")
    myfile.write(output)
    myfile.close()

###################################################

def Unblock(btn):
    output = []
    no_of_imei = 0
    IMEI = clipboard.paste()
    IMEI = str(IMEI)
    IMEI = IMEI.replace("\r\n", ",")
    IMEI = IMEI.split(",")
    del IMEI[-1]

    header = "<spml:batchRequest" + "\n"  \
              "execution"+"="+'"synchronous"' + "\n"  \
              "processing"+"="+'"sequential"' + "\n"  \
              "onError"+"="+'"resume"' + "\n"  \
              "xmlns:spml"+"="+'"urn:siemens:names:prov:gw:SPML:2:0"' + "\n"  \
              "xmlns:nsr"+"="+'"urn:siemens:names:prov:gw:EIR_NSR:1:0"' + "\n"  \
              "xmlns:xsi"+"="+'"http://www.w3.org/2001/XMLSchema-instance">' + "\n"  \
              "<version>EIR_NSR_v30</version>" + "\n"

    initial_data = '<request xsi:type="spml:DeleteRequest">' + "\n"  \
                   '<version>EIR_NSR_v30</version>' + "\n"  \
                   '<objectclass>Device</objectclass>' + "\n"  \
                   '<identifier>'

    Final_data = '</identifier>' + "\n"  \
                 '</request>' + "\n"

    footer = "</spml:batchRequest>"

    output = header

    for imei in IMEI:
        output = output + initial_data + str(IMEI[no_of_imei]) + Final_data
        no_of_imei = no_of_imei + 1

    output = output + footer

    myfile = open("UnBlocking.spml", "w+")
    myfile.write(output)
    myfile.close()




###################################################

###################################################
app = gui("IMEI")
app.addButton("BLOCK", Block,0,1)
app.setButtonWidhts("Block IMEI",2)

app.addButton("UNBLOCK", Unblock,3,1)
app.setButtonWidhts("UnBlock IMEI",2)


app.setAllTextAreaHeights(1)

app.go()