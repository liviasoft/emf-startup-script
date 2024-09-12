#Import relevant python packages
import pyvisa as visa
import datetime
import time

#Setup connection to the MS2090A
rm = visa.ResourceManager()
rm.list_resources()
inst = rm.open_resource("TCPIP::192.168.32.11::9001::SOCKET")
inst.timeout = 50000

inst.read_termination = '\n'
inst.write_termination = '\n'

#Check that the MS2090A is online
print(inst.query("*IDN?"))
time.sleep(2)

#Ensure SPA is active
active_inst = inst.query("INST:SEL?")
print(active_inst)

if active_inst != "SPA":
    inst.write("INST:APPL:STAT HIPM, 0")
    time.sleep(1)
    inst.write("INST:APPL:STAT CAAUSB, 0")
    time.sleep(1)
    inst.write("INST:APPL:STAT SPA, 1")
    time.sleep(30)
    inst.write("INST:SEL SPA")
    time.sleep(60)
    print("SPA selected")

#Preset intrument mode
inst.write("SYST:PRES")
time.sleep(1)
inst.write("SYST:PRES:MODE")
time.sleep(1)

#Set the start and stop frequency; RBW; auto atten for EMF measurement
inst.write("SENS:FREQ:STAR 700 MHZ")
inst.write("SENS:FREQ:STOP 4 GHZ")
time.sleep(0.3)
inst.write("BAND:RES 100 KHz")
inst.write("SENS:FREQ:SWE:TIME 500 MS")
time.sleep(0.3)
inst.write("SENS:POW:RF:ATT:AUTO 1")
inst.write("SENS:POW:RF:GAIN:STAT 0")
time.sleep(0.3)

#Check status of EMF mode and switch to it
emf_state = inst.query("SENS:EMF:STAT?")
print(emf_state)

if emf_state == "0":
    inst.write("SENS:EMF:STAT 1")
    time.sleep(23)    
    
print(inst.query("SENS:EMF:STAT?"))

#Check that logging is enabled
log_state = inst.query("SENS:EMF:LOG?")
print(log_state)

if log_state == "0":
    inst.write("SENS:EMF:LOG 1")
    print(inst.query("SENS:EMF:LOG?"))

#============Start EMF measurement====================

#set emf measurement units
inst.write("UNIT:POW DBM/M2")
time.sleep(0.3)

#set ref level
inst.write("DISP:WIND:TRAC:Y:SCAL:RLEV -5 DBM/M2")
inst.write("DISP:WIND:TRAC:Y:SCAL:RLEV:OFFS 0 DB")

#set measurement time
inst.write("SENS:EMF:MEAS:TIME 10 MIN")
print(inst.query("SENS:EMF:MEAS:TIME?"))
time.sleep(0.3)

#set measurement count
inst.write("SENS:EMF:MEAS:COUN 1")
print(inst.query("SENS:EMF:MEAS:COUN?"))
time.sleep(0.3)

#set emf dwell time
inst.write("SENS:EMF:AXIS:TIME 1 S")
time.sleep(0.3)

#set ICNIRP limit
print(inst.query("SENS:EMF:ICN?"))
inst.write("SENS:EMF:ICN:STAT 1")
inst.write("SENS:EMF:ICN:STAT?")
time.sleep(1)

i = 0

while i < 24:
    inst.write("SENS:EMF:RUN 1")
    i = i + 1
    print(i)
    time.sleep(620)   
            
#Disable EMF mode
inst.write("SENS:EMF:STAT 0")
time.sleep(23)
print(inst.query("SENS:EMF:STAT?"))

#Preset intrument mode
inst.write("SYST:PRES")
inst.write("SYST:PRES:MODE")