# -*- coding: utf-8 -*-
print("EUSignCP Python ASiC Signer Test:")

from EUSignCP import *

privateKeyFilePath = b"Key-6.dat"
privateKeyPassword = b"12345677"
file = open(privateKeyFilePath, "rb")
pkData = file.read()
file.close()

signAlgo = EU_CTX_SIGN_ECDSA_WITH_SHA
asicType = EU_ASIC_TYPE_E
referencesCount = (2, 1)[asicType == EU_ASIC_TYPE_S]
signFileExt = (".asice", ".asics")[asicType == EU_ASIC_TYPE_S]

EULoad()
pIface = EUGetInterface()
try:
	pIface.Initialize()
except Exception as e:
	print ("Initialize failed "  + str(e))
	EUUnload()
	exit()

print("Library Initialized")

libContext = []
pkContext = []

try:
	pIface.CtxCreate(libContext)
except Exception as e:
	dError = eval(str(e))
	print ("CtxCreate failed "  + str(dError['ErrorCode']) + ". Description: " + dError['ErrorDesc'])
	pIface.Finalize()
	EUUnload()
	exit()

ownerInfo = {}

try:
	pIface.CtxReadPrivateKeyBinary(libContext[0], pkData, len(pkData), privateKeyPassword, pkContext, ownerInfo)
except Exception as e:
	dError = eval(str(e))
	print ("Key reading failed "  + str(dError['ErrorCode']) + ". Description: " + dError['ErrorDesc'])
	pIface.CtxFree(libContext[0])
	pIface.Finalize()
	EUUnload()
	exit()

pRef = ["1.txt"]
pData = [b"1234"]
if asicType == EU_ASIC_TYPE_E:
	pRef.append("2.txt")
	pData.append(b"abcdef")

pASiCData = []

try:
	pIface.ASiCCreateEmptySign(asicType, EU_ASIC_SIGN_TYPE_CADES, pRef, pData, pASiCData)
except Exception as e:
	dError = eval(str(e))
	print ("ASiCCreateEmptySign error "  + str(dError['ErrorCode']) + ". Description: " + dError['ErrorDesc'])
	pIface.CtxFreePrivateKey(pkContext[0])
	pIface.CtxFree(libContext[0])
	pIface.Finalize()
	EUUnload()
	exit()

pSignRef = []
pAttrsHash = []
pASiCSign = []

try:
	pIface.ASiCCreateSignerBegin(signAlgo, asicType, EU_ASIC_SIGN_TYPE_CADES, pRef, pASiCData[0], len(pASiCData[0]), pSignRef, pAttrsHash, pASiCSign)
except Exception as e:
	dError = eval(str(e))
	print ("ASiCCreateSignerBegin error "  + str(dError['ErrorCode']) + ". Description: " + dError['ErrorDesc'])
	pIface.CtxFreePrivateKey(pkContext[0])
	pIface.CtxFree(libContext[0])
	pIface.Finalize()
	EUUnload()
	exit()

pSign = []

try:
	pIface.CtxSignHashValue(pkContext[0], signAlgo, pAttrsHash[0], len(pAttrsHash[0]), True, pSign)
except Exception as e:
	dError = eval(str(e))
	print ("CtxSignHashValue error "  + str(dError['ErrorCode']) + ". Description: " + dError['ErrorDesc'])
	pIface.CtxFreePrivateKey(pkContext[0])
	pIface.CtxFree(libContext[0])
	pIface.Finalize()
	EUUnload()
	exit()

pIface.CtxFreePrivateKey(pkContext[0])
pIface.CtxFree(libContext[0])

try:
	pIface.ASiCCreateSignerEnd(asicType, EU_ASIC_SIGN_TYPE_CADES, EU_SIGN_TYPE_CADES_X_LONG, pSignRef[0], pSign[0], len(pSign[0]), pASiCSign[0], len(pASiCSign[0]), pASiCData)
except Exception as e:
	dError = eval(str(e))
	print ("ASiCCreateSignerEnd error "  + str(dError['ErrorCode']) + ". Description: " + dError['ErrorDesc'])
	pIface.Finalize()
	EUUnload()
	exit()

file = open("Sign" + signFileExt, "wb")
file.write(pASiCData[0])
file.close()

pIface.Finalize()
EUUnload()

print("EUSignCP Python ASiC Signer Test done.")
