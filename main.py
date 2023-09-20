# This script conducts curtailment analysis within SC2 Boundary over a period of 1 year
# The profiles (data points) used for interconnectors, wind driven generators and loads are artificial
# The output is a 2-dimension table which records the amount of overload of each asset in every hour over the period

from ipsa import *
from untilities import *
import numpy as np
import pandas as pd

# Initialise ipsa and get network
ipsasys = IscInterface()
network = ipsasys.ReadFile(r"E:\Documents\IPSA_2_(internship)\IPSA_Files\Network SC2 Boundary.i2f")
# -------------------------------------------------------------------------------------------------
# Specify the length of time period in curtailment analysis
dTime_Days = 30

# Get network components
dicSynMachines = network.GetSynMachines()

dicLoads = network.GetLoads()

dicBranches = network.GetBranches()
lBranchesName = []
for strBranchName in dicBranches.keys():
    lBranchesName.append(strBranchName)

dicTransformers = network.GetTransformers()
lTransformersName = []
for strTransfName in dicTransformers.keys():
    lTransformersName.append(strTransfName)

# Get name of assets in a single list
lAssetsName = lBranchesName + lTransformersName

# -------------------------------------------------------------------------------------------------
# Get name of synchronous machines and loads
lSynMachinesName = []
lLoadsName = []
for strSynMachineName in dicSynMachines.keys():
    lSynMachinesName.append(strSynMachineName)
    print(strSynMachineName)

for strLoadName in dicLoads.keys():
    lLoadsName.append(strLoadName)
    print(strLoadName)

# -------------------------------------------------------------------------------------------------
# Generate artificial profiles of generations
lGeneratorsName = ['Rampion_SUB.Rampion', 'London Array_SUB.London Array', 'E de F.IFA To France', 'Eleclink.Eleclink To France',
                   'NEMO.NEMO To Belgium', 'Thanet_SUB.Thanet']
lGeneratorsMeanPowerMW = [400, 800, 800, -800, -300, 600]
lGeneratorsPowerAmplitudeMW = [400, 130, 300, 600, 400, 300]
lGeneratorsPowerFrequency = [10, 10, 36, 36, 36, 10]
lGeneratorsPowerOffsetRadian = [0, 0.5, 1, 1.5, 2, 2.5]

llGenerationProfiles = []
for i in range(len(lGeneratorsName)):
    lx, ly = GenerateSinusoidalProfile(lGeneratorsMeanPowerMW[i], lGeneratorsPowerAmplitudeMW[i],
                                       lGeneratorsPowerFrequency[i], lGeneratorsPowerOffsetRadian[i], dTime_Days*24)
    llGenerationProfiles.append(ly)

dfGenerationProfiles = pd.DataFrame()
for i in range(len(lGeneratorsName)):
    str_temp = lGeneratorsName[i]
    dfGenerationProfiles[str_temp] = llGenerationProfiles[i]

# Generate artificial profiles for loads
lLoadsName = ['BOLN_SUB.Load', 'CANTN_SUB.Load', 'Ninfield.Load',
              'Richborough.Load', 'Sellindge.Load']
lLoadsMeanPowerMW = [500, 600, 300, 400, 200]
lLoadsPowerAmplitudeMW = [200, 300, 200, 250, 200]
lLoadsPowerFrequency = [36, 36, 36, 36, 36]
lLoadsPowerOffsetRadian = [1, 2, 0.5, 1.5, 0]

llLoadProfiles = []
for i in range(len(lLoadsName)):
    lx, ly = GenerateSinusoidalProfile(lLoadsMeanPowerMW[i], lLoadsPowerAmplitudeMW[i],
                                       lLoadsPowerFrequency[i], lLoadsPowerOffsetRadian[i], dTime_Days*24)
    llLoadProfiles.append(ly)

dfLoadProfiles = pd.DataFrame()
for i in range(len(lLoadsName)):
    str_temp = lLoadsName[i]
    dfLoadProfiles[str_temp] = llLoadProfiles[i]

writer = pd.ExcelWriter(r"Excel Files\results.xlsx", engine='openpyxl')
dfGenerationProfiles.to_excel(writer, sheet_name='Generation Profile')
dfLoadProfiles.to_excel(writer, sheet_name='Load Profile')
# -------------------------------------------------------------------------------------------------
# Set up a dataframe to store results
dfResults = pd.DataFrame({'Asset Name': lAssetsName})
# Run load flow at each hour and store overload results
for i in range(dTime_Days*24):
    # Set the generation and load values
    for j in range(6):
        SetSynMachineGenerationMW(dicSynMachines, lGeneratorsName[j], dfGenerationProfiles.iloc[i, j])
    for j in range(6):
        SetLoadRealPowerMW(dicLoads, lLoadsName[j], dfLoadProfiles.iloc[i, j])
    # Do load flow
    network.DoDCLoadFlow()
    # Get load flow results (percentage overload of branches and transformers)
    dfResults[str(i)] = GetAssetsOverloadPercentage(dicBranches, dicTransformers, 0)

# Output the result in an Excel file
dfResults.to_excel(writer, sheet_name='Raw Data')

# -------------------------------------------------------------------------------------------------
# Go through the raw data column by column to check if any asset is overloaded
lOverloadHour = []
for i in range(dTime_Days*24):
    for j in range(len(lAssetsName)):
        if dfResults.iloc[j, i+1] > 1:
            lOverloadHour.append(i)

# Output the data of overloaded hours in second sheet
dfResults_OverLoaded = pd.DataFrame({'Asset Name': lAssetsName})
for i in lOverloadHour:
    dfResults_OverLoaded[str(i)] = dfResults[str(i)]
dfResults_OverLoaded.to_excel(writer, sheet_name='Overloaded Data')

# Define to what percentage the generator will be curtailed each time
dCurtailStep = 0.1
# Create a dataframe to store generation values after curtailment
dfResults_AfterCurt = pd.DataFrame({'Generator Name': lGeneratorsName})
# Create a dafaframe to store generation values both before and after curtailment
dfResults_BeforeAndAfterCurt = pd.DataFrame({'Generator Name': lGeneratorsName})
# Go through the hours in which overload happens and do curtailment until overloading disappears
for i in lOverloadHour:
    for j in range(len(lGeneratorsName)):
        # Set a boolean to indicate if overloading has been overcome
        bFlag = 0
        dCurtailPercentage = 0
        while bFlag == 0 and dCurtailPercentage < 1:
            # Set values of generators
            lCurrentGenerationProfile = dfGenerationProfiles.iloc[i, :].tolist()
            SetSynMachinesGenerationMW(dicSynMachines, lGeneratorsName, lCurrentGenerationProfile)
            # Set values of loads
            lCurrentLoadProfile = dfLoadProfiles.iloc[i, :].tolist()
            SetLoadsRealPowerMW(dicLoads, lLoadsName, lCurrentLoadProfile)

            # Increase curtailment percentage and ensure it won't exceed 1
            dCurtailPercentage += dCurtailStep
            if dCurtailPercentage > 1:
                dCurtailPercentage = 1
            # Curtail first generator by 10% of its generation
            CurtailGenerationByPercentage(dicSynMachines, lGeneratorsName[j],
                                          dfGenerationProfiles.iloc[i, j], dCurtailPercentage)

            # Do DC load flow and get assets overload percentage
            network.DoDCLoadFlow()
            lCurrent_Result = GetAssetsOverloadPercentage(dicBranches, dicTransformers, 0)
            # Go through the overload percentage and check if any asset is overloaded
            for r in lCurrent_Result:
                if r > 1:
                    bFlag = 1
                    break
    lGenerationsMW_AfterCurt = GetSMGenerationsMW(dicSynMachines, lGeneratorsName)
    # Store data
    dfResults_AfterCurt[str(i)] = lGenerationsMW_AfterCurt
    dfResults_BeforeAndAfterCurt[str(i)] = dfGenerationProfiles.iloc[i, :].tolist()
    dfResults_BeforeAndAfterCurt[str(i) + '_after'] = lGenerationsMW_AfterCurt

# Output generation values after curtailment at each hour
dfResults_AfterCurt.to_excel(writer, sheet_name='Generation After Curtailment')
dfResults_BeforeAndAfterCurt.to_excel(writer, sheet_name='Generation Comparison')
writer.close()





