
from ipsa import *
import numpy as np
import pandas as pd

def SetSynMachineGenerationMW(dicSynMachines, strSynMachineName, dGenerationMW):
    bFlag = False
    for strSMName, objSM in dicSynMachines.items():
        if strSMName == strSynMachineName:
            objSM.SetDValue(9, dGenerationMW)
            bFlag = True
    return bFlag

def SetLoadRealPowerMW(dicLoads, strLoadName, dLoadRealPowerMW):
    bFlag = False
    for strLName, objL in dicLoads.items():
        if strLName == strLoadName:
            objL.SetDValue(6, dLoadRealPowerMW)
            bFlag = True
    return bFlag

def SetSynMachinesGenerationMW(dicSynMachines, lSynMachinesName, lSynMachinesGenerationMW):
    bFlag = True
    cnt = 0
    for strSMName, objSM in dicSynMachines.items():
        if strSMName == lSynMachinesName[cnt]:
            objSM.SetDValue(9, lSynMachinesGenerationMW[cnt])
        else:
            bFlag = False
    return bFlag

def SetLoadsRealPowerMW(dicLoads, lLoadsName, lLoadsRealPowerMW):
    bFlag = True
    cnt = 0
    for strLName, objL in dicLoads.items():
        if strLName == lLoadsName[cnt]:
            objL.SetDValue(6, lLoadsRealPowerMW[cnt])
        else:
            bFlag = False
    return bFlag

def GenerateSinusoidalProfile(dMean, dAmplitude, dFrequency, dOffsetRadian, dN_Points):
    lx = []
    ly = []
    cnt = 0
    for i in np.linspace(0, 2*np.pi*dFrequency, dN_Points):
        lx.append(cnt)
        ly.append(np.sin(i+dOffsetRadian) * dAmplitude + dMean)
    return lx, ly

def GetAssetsOverloadPercentage(dicBranches, dicTransformers, nRatingIndex):
    lBranchOverloadPercentage = []
    for strBranchName, objBranch in dicBranches.items():
        lBranchOverloadPercentage.append(abs(objBranch.GetDCLFSendRealPowerMW()) / objBranch.GetRatingMVA(nRatingIndex))

    lTransfOverloadPercentage = []
    for strTransfName, objTransf in dicTransformers.items():
        lTransfOverloadPercentage.append(abs(objTransf.GetDCLFSendRealPowerMW()) / objTransf.GetRatingMVA(nRatingIndex))

    lAssetsOverloadPercentage = lBranchOverloadPercentage + lTransfOverloadPercentage
    return lAssetsOverloadPercentage

def CurtailGenerationByPercentage(dicSynMachines, strSynMachineName, dReferenceGenerationMW, dCurtailmentPercentage):
    bFlag = False
    for strSMName, objSM in dicSynMachines.items():
        if strSMName == strSynMachineName:
            objSM.SetDValue(9, dReferenceGenerationMW * (1 - dCurtailmentPercentage))
            bFlag = True
    return bFlag

def CurtailLoadByPercentage(dicLoads, strLoadName, dReferenceLoadMW, dCurtailmentPercentage):
    bFlag = False
    for strLName, objL in dicLoads.items():
        if strLName == strLoadName:
            objL.SetDValue(9, dReferenceLoadMW * (1 - dCurtailmentPercentage))
            bFlag = True
    return bFlag

def GetSMGenerationsMW(dicSynMachines, lGeneratorsName):
    lGenerationsMW = []
    for strSMName, objSM in dicSynMachines.items():
        if strSMName in lGeneratorsName:
            lGenerationsMW.append(objSM.GetDValue(9))
    return lGenerationsMW


