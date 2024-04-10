# -*- coding: utf-8 -*-
print("EUSignCP Python ASiC-E Sign Test:")

from EUSignCP import *

privateKeyFilePath = b"pb_2836413030.jks"
privateKeyPassword = b"2808Andru77"

EULoad()
pIface = EUGetInterface()
try:
	pIface.Initialize()
except Exception as e:
	print ("Initialize failed"  + str(e))
	EUUnload()
	exit()

print("Library Initialized")

try:
	if not pIface.IsPrivateKeyReaded():
		pIface.ReadPrivateKeyFile(privateKeyFilePath, privateKeyPassword, None)
except Exception as e:
	dError = eval(str(e))
	print ("Key reading failed"  + str(dError['ErrorCode']) + ". Description: " + dError['ErrorDesc'])
	pIface.Finalize()
	EUUnload()
	exit()

if pIface.IsPrivateKeyReaded():
	print("Key success read")
else:
	print("Key read failed")
	pIface.Finalize()
	EUUnload()
	exit()

pRef = ["data1", "data2"]
pData = [b"1234", b"abcd"]
pASiCData = []

try:
	pIface.ASiCSignData(EU_ASIC_TYPE_E, EU_ASIC_SIGN_TYPE_CADES, EU_ASIC_SIGN_LEVEL_BES, pRef, pData, pASiCData)
except Exception as e:
	dError = eval(str(e))
	print ("Sign error"  + str(dError['ErrorCode']) + ". Description: " + dError['ErrorDesc'])
	pIface.Finalize()
	EUUnload()
	exit()

try:
	pIface.ASiCVerifyData(0, pASiCData[0], len(pASiCData[0]), None)
except Exception as e:
	dError = eval(str(e))
	print ("Verify error"  + str(dError['ErrorCode']) + ". Description: " + dError['ErrorDesc'])
	pIface.Finalize()
	EUUnload()
	exit()

pRefOut = []

try:
	pIface.ASiCGetSignReferences(0, pASiCData[0], len(pASiCData[0]), pRefOut)
except Exception as e:
	dError = eval(str(e))
	print ("Verify error"  + str(dError['ErrorCode']) + ". Description: " + dError['ErrorDesc'])
	pIface.Finalize()
	EUUnload()
	exit()

if pRefOut[0] != pRef[0] or pRefOut[1] != pRef[1]:
	print ("Returned wrong references")
	pIface.Finalize()
	EUUnload()
	exit()

pDataOut = []

try:
	pIface.ASiCGetReference(pASiCData[0], len(pASiCData[0]), pRefOut[1], pDataOut)
except Exception as e:
	dError = eval(str(e))
	print ("Verify error"  + str(dError['ErrorCode']) + ". Description: " + dError['ErrorDesc'])
	pIface.Finalize()
	EUUnload()
	exit()

if pDataOut[0] != pData[1]:
	print ("Returned wrong data")
	pIface.Finalize()
	EUUnload()
	exit()

pIface.Finalize()
EUUnload()

print("EUSignCP Python ASiC-E Sign Test done.")
