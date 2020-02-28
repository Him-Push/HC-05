import wmi    
c=wmi.WMI()
global m
wql="Select DeviceID from Win32_PnPEntity where name='HC-05'"
wql1="Select deviceid from win32_SerialPort"
global l1
l=[]
l1=[]
l3=[]

for item in c.query(wql):
    s=str(item)
for item1 in c.query(wql1):
    l.append(item1)
serial1=str(l)
t=serial1.split('="')
for i in range(1,len(t)):
    serial3=str(t[i])
    l1.append(serial3[0:5])
#print(l1)

x=s.find("BLUETOOTHDEVICE")
s=s[x+16:]

t=[]
for i in range(0,len(s)):
    if(s[i]=='"'):
        break
    else:
        t.append(s[i])

str1=""
m=str1.join(t)
print(m)
wql3="Select PNPDeviceID from win32_SerialPort where deviceid='"
for i in range(0,len(l1)):
    
    wql4=wql3+l1[i]+str("'")
    l3.append(wql4)
for i in range(0,len(l3)):
    
    wql5=str(l3[i])
    
    for item7 in c.query(wql5):
        
        sfinder=str(item7)
        
        x2=sfinder.find(m)
        
        if(x2>1):
            print("sucess")
            global j
            j=i
            #print(j)
            
print(l1[j])


    
