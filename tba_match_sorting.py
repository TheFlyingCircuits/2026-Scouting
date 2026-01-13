from dataclasses import dataclass
import os
import json
import shared_classes
from typing import List

tbaFileName = 'TBAMatches.json'

global tbaMatchesWithPlayoffs
global tbaQualsMatches
global scoutedQualsMatchesRed
global scoutedQualsMatchesBlue
tbaQualsMatches = []
scoutedQualsMatchesRed = []
scoutedQualsMatchesBlue = []

autoL1PointsValue = 3
autoL2PointsValue = 4
autoL3PointsValue = 6
autoL4PointsValue = 7
L1PointsValue = 2
L2PointsValue = 3
L3PointsValue = 4
L4PointsValue = 5
processorPointsValue = 2
netPointsValue = 4


def makeBothAllianceMatchClass(SingleTeamSingleMatchEntrysList: List[shared_classes.SingleTeamSingleMatchEntry], all_team_data: dict[int, shared_classes.TeamData]) -> shared_classes.scoutingAccuracyMatch:
    outputData = shared_classes.scoutingAccuracyMatch
    redTeamNumbs = []
    blueTeamNumbs = []
    # random singleteammatchentrys because I want .team_num and .auto_points to show up for convinince
    redAllianceInOrderData = [SingleTeamSingleMatchEntrysList[0],SingleTeamSingleMatchEntrysList[1],SingleTeamSingleMatchEntrysList[2]]
    blueAllianceInOrderData = [SingleTeamSingleMatchEntrysList[0],SingleTeamSingleMatchEntrysList[1],SingleTeamSingleMatchEntrysList[2]]
    matchData = tbaQualsMatches[SingleTeamSingleMatchEntrysList[0].qual_match_num - 1]
    for i, teamString in enumerate(matchData["alliances"]["blue"]["team_keys"]):
        teamNumWithoutString = int(teamString.replace("frc", ""))
        blueTeamNumbs.append(teamNumWithoutString)
    for i, teamString in enumerate(matchData["alliances"]["red"]["team_keys"]):
        teamNumWithoutString = int(teamString.replace("frc", ""))
        redTeamNumbs.insert(i,teamNumWithoutString)
    for teamMatchData in SingleTeamSingleMatchEntrysList:
        teamNum = teamMatchData.team_num
        for i in range(6):
            if i <= 2:
                if teamNum == redTeamNumbs[i]:
                    redAllianceInOrderData[i] = teamMatchData
            else:
                if teamNum == blueTeamNumbs[i-3]:
                    blueAllianceInOrderData[i-3] = teamMatchData
    # Validating the data now

    redScoutersQuantitativeInacuracy = [0.0,0.0,0.0]
    redScoutersMissedClimbs = [False,False,False]
    blueScoutersQuantitativeInacuracy = [0.0,0.0,0.0]
    blueScoutersMissedClimbs = [False,False,False]

    # auto
    for i in range (2):
        allianceList = []
        alliance = ""
        autoOffPercent = 0.0
        teleOffPercent = 0.0
        endGameOffPercent = 0.0
        if i == 0:
            allianceList = redAllianceInOrderData
            alliance = "red"
        else:
            allianceList = blueAllianceInOrderData
            alliance = "blue"
    
        # print(autoOffPercent)
        scoutedCoralL2TotalAuto = allianceList[0].autoL2 + allianceList[1].autoL2 + allianceList[2].autoL2
        scoutedCoralL3TotalAuto = allianceList[0].autoL3 + allianceList[1].autoL3 + allianceList[2].autoL3
        scoutedCoralL4TotalAuto = allianceList[0].autoL4 + allianceList[1].autoL4 + allianceList[2].autoL4

        tbaL2Auto =  matchData["score_breakdown"][alliance]["autoReef"]["tba_botRowCount"]
        tbaL3Auto =  matchData["score_breakdown"][alliance]["autoReef"]["tba_midRowCount"]
        tbaL4Auto =  matchData["score_breakdown"][alliance]["autoReef"]["tba_topRowCount"]
        
        autoOffPercent += abs(scoutedCoralL2TotalAuto - tbaL2Auto) * 0.3
        autoOffPercent += abs(scoutedCoralL3TotalAuto - tbaL3Auto) * 0.3
        autoOffPercent += abs(scoutedCoralL4TotalAuto - tbaL4Auto) * 0.3

        for i in range(3):
            teamNum = allianceList[i].team_num
            allScoutedAutoPoints = (allianceList[i].autoL1*autoL1PointsValue) + (allianceList[i].autoL2*autoL2PointsValue) 
            + (allianceList[i].autoL3*autoL3PointsValue) + (allianceList[i].autoL4*autoL4PointsValue) 
            + (allianceList[i].autoNet*netPointsValue) + (allianceList[i].autoProcessor*processorPointsValue) 

            aveAllAutoPoints = all_team_data[teamNum].aveAutoL1Points+all_team_data[teamNum].aveAutoL2Points
            +all_team_data[teamNum].aveAutoL3Points+all_team_data[teamNum].aveAutoL4Points+all_team_data[teamNum].aveAutoNetPoints
            +all_team_data[teamNum].aveAutoProcessorPoints

            predictedDataScewAuto = 0.0
            if allScoutedAutoPoints == 0:
                predictedDataScewAuto = aveAllAutoPoints * 1.0 * autoOffPercent
                # predictedDataScewAuto = 0
            elif aveAllAutoPoints == 0:
                predictedDataScewAuto = allScoutedAutoPoints * 1.0 * autoOffPercent
                # predictedDataScewAuto = 0
            elif allScoutedAutoPoints + aveAllAutoPoints != 0:
                predictedDataScewAuto = (0.15 * allScoutedAutoPoints/aveAllAutoPoints) * autoOffPercent
                # predictedDataScewAuto = 0
            if alliance == "red":
                redScoutersQuantitativeInacuracy[i] += predictedDataScewAuto
            else:
                blueScoutersQuantitativeInacuracy[i] += predictedDataScewAuto
            leavesOff = 0.0
        for i in range(3):
            teamAutoLineStr = matchData["score_breakdown"][alliance]["autoLineRobot"+str(i+1)]
            teamAutoLineBool = teamAutoLineStr == "Yes"
            if allianceList[i].leave != teamAutoLineBool:
                if alliance == "red":
                    redScoutersQuantitativeInacuracy[i] += 0.4 * (1.0 / 3.0)
                else:
                    blueScoutersQuantitativeInacuracy[i] += 0.4 * (1.0 / 3.0) #this is about 13%
                leavesOff += 1.0
        autoOffPercent += 0.4 * (leavesOff / 3.0) # max if can be is 0.3 or 30%
        # print(abs(scoutedCoralL4TotalAuto - tbaL4Auto) * 0.3)
        # print(str(tbaL4Auto) + " real "+ str(tbaL4Auto) + " "+ alliance + " "+ str(allianceList[0].qual_match_num) + 
        #       " " + str(matchData["match_number"]))
        # print(autoOffPercent)

        # Teleop calculation

        scoutedCoralL1Total = allianceList[0].autoL1 + allianceList[0].teleL1 
        + allianceList[1].autoL1 + allianceList[1].teleL1 + allianceList[2].autoL1 +allianceList[2].teleL1
        scoutedCoralL2TotalTele = allianceList[0].teleL2 + allianceList[1].teleL2 + allianceList[2].teleL2
        scoutedCoralL3TotalTele = allianceList[0].teleL3 + allianceList[1].teleL3 + allianceList[2].teleL3
        scoutedCoralL4TotalTele = allianceList[0].teleL4 + allianceList[1].teleL4 + allianceList[2].teleL4

        scoutedNetAlgae = allianceList[0].autoNet + allianceList[0].teleNet 
        + allianceList[1].autoNet + allianceList[1].teleNet + allianceList[2].autoNet +allianceList[2].teleNet
        scoutedProcessorAlgae = allianceList[0].autoProcessor + allianceList[0].teleProcessor 
        + allianceList[1].autoProcessor + allianceList[1].teleProcessor + allianceList[2].autoProcessor +allianceList[2].teleProcessor

        scoutedNetAlgae = scoutedNetAlgae - scoutedProcessorAlgae

        tbaL1 =  matchData["score_breakdown"][alliance]["teleopReef"]["trough"]
        tbaL2Tele =  matchData["score_breakdown"][alliance]["teleopReef"]["tba_botRowCount"]
        tbaL3Tele =  matchData["score_breakdown"][alliance]["teleopReef"]["tba_midRowCount"]
        tbaL4Tele =  matchData["score_breakdown"][alliance]["teleopReef"]["tba_topRowCount"]

        tbaNetAlgae =  matchData["score_breakdown"][alliance]["netAlgaeCount"]
        tbaProccesorAlgae =  matchData["score_breakdown"][alliance]["wallAlgaeCount"]

        teleOffPercent += abs(scoutedCoralL1Total - tbaL1) * 0.05
        teleOffPercent += abs(scoutedCoralL2TotalTele - tbaL2Tele) * 0.05
        teleOffPercent += abs(scoutedCoralL3TotalTele - tbaL3Tele) * 0.05
        teleOffPercent += abs(scoutedCoralL4TotalTele - tbaL4Tele) * 0.05
        teleOffPercent += abs(scoutedProcessorAlgae - tbaProccesorAlgae) * 0.05
        # make it so the more processor then we trust the total net count less bc they just count how many are in the net
        howMuchWeTrustNetCount = 0.1
        if tbaProccesorAlgae > 0:
            if tbaProccesorAlgae == 1:
                howMuchWeTrustNetCount = 0.05
            else:
                howMuchWeTrustNetCount = 0.1 *(1/tbaProccesorAlgae)
        if(tbaNetAlgae <= 0 and tbaProccesorAlgae > 0):
            howMuchWeTrustNetCount = 0
            # print("kidna sketchy")
        # teleOffPercent += abs(scoutedNetAlgae - tbaNetAlgae) * howMuchWeTrustNetCount
        # print(abs(scoutedNetAlgae - tbaNetAlgae) * howMuchWeTrustNetCount)
        # print(teleOffPercent)

        for i in range(3):
            teamNum = allianceList[i].team_num
            allScoutedTelePoints = (allianceList[i].teleL1*L1PointsValue) + (allianceList[i].teleL2*L2PointsValue) 
            + (allianceList[i].teleL3*L3PointsValue) + (allianceList[i].teleL4*L4PointsValue) 
            + (allianceList[i].teleNet*netPointsValue) + (allianceList[i].teleProcessor*processorPointsValue) 

            aveAllTelePoints = all_team_data[teamNum].aveTeleL1Points+all_team_data[teamNum].aveTeleL2Points
            +all_team_data[teamNum].aveTeleL3Points+all_team_data[teamNum].aveTeleL4Points+all_team_data[teamNum].aveTeleNetPoints
            +all_team_data[teamNum].aveTeleProcessorPoints

            predictedDataScewTele = 0.0
            if allScoutedTelePoints == 0:
                predictedDataScewTele = aveAllTelePoints * 1.0 * teleOffPercent
                # predictedDataScewAuto = 0
            elif aveAllTelePoints == 0:
                predictedDataScewTele = allScoutedTelePoints * 1.0 * teleOffPercent
                # predictedDataScewAuto = 0
            elif allScoutedTelePoints + aveAllTelePoints != 0:
                predictedDataScewTele = (0.15 * allScoutedTelePoints/aveAllTelePoints) * teleOffPercent
                # predictedDataScewAuto = 0
            if alliance == "red":
                # if i != 2:
                redScoutersQuantitativeInacuracy[i] += predictedDataScewTele
            else:
                # if i != 2:
                blueScoutersQuantitativeInacuracy[i] += predictedDataScewTele
        
        climbsOff = 0.0
        for i in range(3):
            teamClimbStr = matchData["score_breakdown"][alliance]["endGameRobot"+str(i+1)]
            match teamClimbStr:
                case "None":
                    teamClimbStr = "None of the above"
                case "Parked":
                    teamClimbStr = "Park in the barge zone"
                case "ShallowCage":
                    teamClimbStr = "Climb on the shallow cage"
                case "DeepCage":
                    teamClimbStr = "Climb on the deep cage"
            if allianceList[i].climb != teamClimbStr:
                if alliance == "red":
                    redScoutersQuantitativeInacuracy[i] += 1.0 * (1.0 / 3.0)
                    redScoutersMissedClimbs[i] = True
                    # print(redScoutersMissedClimbs[i])
                else:
                    blueScoutersQuantitativeInacuracy[i] += 1.0 * (1.0 / 3.0)
                    blueScoutersMissedClimbs[i] = True
                climbsOff += 1.0
            # print(allianceList[i].climb+" "+teamAutoLineStr)
            # print("sf")
        endGameOffPercent += 1.0 * (climbsOff / 3.0)
        # print(str(endGameOffPercent) + alliance + " "+ str(allianceList[0].qual_match_num) + " " + str(matchData["match_number"]))
        # if (autoOffPercent+teleOffPercent+endGameOffPercent == 0.0):
        #     print(autoOffPercent+teleOffPercent+endGameOffPercent)
        autoOffPercent = autoOffPercent*100
        teleOffPercent = teleOffPercent*100
        endGameOffPercent = endGameOffPercent*100
        if alliance == "red":
            outputData.overallInaccuracyRed = autoOffPercent+teleOffPercent+endGameOffPercent
            outputData.autoInaccuracyRed = autoOffPercent
            outputData.teleInaccuracyRed = teleOffPercent
            outputData.endGameInaccuracyRed = endGameOffPercent
            outputData.matchNumRed = matchData["match_number"]
            outputData.scouterOneNameRed = allianceList[0].commenter
            outputData.scouterTwoNameRed = allianceList[1].commenter
            outputData.scouterThreeNameRed = allianceList[2].commenter
            outputData.scouterOneInacuracyRed = redScoutersQuantitativeInacuracy[0] * 100.0
            outputData.scouterTwoInacuracyRed = redScoutersQuantitativeInacuracy[1] * 100.0
            outputData.scouterThreeInacuracyRed= redScoutersQuantitativeInacuracy[2] * 100.0
            outputData.scouterOneMissClimbRed= redScoutersMissedClimbs[0]
            outputData.scouterTwoMissClimbRed= redScoutersMissedClimbs[1]
            outputData.scouterThreeMissClimbRed= redScoutersMissedClimbs[2]
        else:
            outputData.overallInaccuracyBlue = autoOffPercent+teleOffPercent+endGameOffPercent
            outputData.autoInaccuracyBlue = autoOffPercent
            outputData.teleInaccuracyBlue = teleOffPercent
            outputData.endGameInaccuracyBlue = endGameOffPercent
            outputData.matchNumBlue = matchData["match_number"]
            outputData.scouterOneNameBlue = allianceList[0].commenter
            outputData.scouterTwoNameBlue = allianceList[1].commenter
            outputData.scouterThreeNameBlue = allianceList[2].commenter
            outputData.scouterOneInacuracyBlue = blueScoutersQuantitativeInacuracy[0] * 100.0
            outputData.scouterTwoInacuracyBlue = blueScoutersQuantitativeInacuracy[1] * 100.0
            outputData.scouterThreeInacuracyBlue = blueScoutersQuantitativeInacuracy[2] * 100.0
            outputData.scouterOneMissClimbBlue= blueScoutersMissedClimbs[0]
            outputData.scouterTwoMissClimbBlue= blueScoutersMissedClimbs[1]
            outputData.scouterThreeMissClimbBlue= blueScoutersMissedClimbs[2]
    # print(redScoutersQuantitativeInacuracy[0])
    return outputData   



        
def initializeTBAData():
    try:
        with open(tbaFileName, 'r') as file:
            tbaMatchesWithPlayoffs = json.load(file)
        #print(tbaMatches)
    except FileNotFoundError:
        print("Error: The file was not found.")
    except json.JSONDecodeError:
        print("Error: Failed to decode JSON from the file.")

    for i, match in enumerate(tbaMatchesWithPlayoffs):
        if match["comp_level"] == "qm":
            tbaQualsMatches.insert(match["match_number"] - 1, match)
            # print(match["match_number"])
        # print(match["comp_level"])

#print(tbaQualsMatches[11]["match_number"])

