import csv
import os
import xlsxwriter
import tba_match_sorting
import shared_classes
import json
from dataclasses import dataclass, asdict, is_dataclass

input_file_name = "input.csv"
output_file_name = "output_data.xlsx"

allMatchesValidationData = [shared_classes.scoutingAccuracyMatch]
all_team_match_entries = []
all_team_data: dict[int, shared_classes.TeamData] = {}

team_num_list = []
output_worksheets = []

total_points_rankings = []
tele_score_rankings = []
auto_score_rankings = []
rice_score_rankings = []
total_points_rankings_team_names = []
tele_score_rankings_team_names = []
auto_score_rankings_team_names = []
rice_score_rankings_team_names = []
rice_scores = []

DATA_START_ROW = 37
AVERAGES_ROW = 33
STATISTICS_START_ROW = 16
STATISTICS_START_COL = 0
CHART_START_ROW = 1
CHART_ROW_SPACING = 16
FIRST_CHART_COL = "A"
SECOND_CHART_COL = "E"
THIRD_CHART_COL = "I"
NUM_OF_TOP_TEAMS_TO_COLOR_PER_CATEGORY = 5
current_match_count = 0

fuelPointsValue = 1
autoL1ClimbPointsValue = 15
l1ClimbPointsValue = 10
l1ClimbPointsValue = 20
l1ClimbPointsValue = 30

chart_colors = {
    "RED": "#EA5545",
    "BLUE": "#27AEEF",
    "YELLOW": "#FFCA3A",
    "GREEN": "#4DA167",
    "PURPLE": "#805D93",
    "BLACK": "#252323",
}


def parse_team_number(num):
    # The team number could come in as a poorly formatted string or a number, this function
    # helps standardize the input
    if type(num) == int:
        parsed_num = num
    elif type(num) == str:
        if num != "":
            parsed_num = int(float(num))
        else:
            parsed_num = -1

    return parsed_num


def parse_match_number(num):
    # The match number could come in as a poorly formatted string or a number, this function
    # helps standardize the input
    if type(num) == int:
        parsed_num = num
    elif type(num) == str:
        if num != "":
            parsed_num = int(float(num))
        else:
            parsed_num = -1

    return parsed_num


def get_highest_number(numbers):
    # Used because answers can be "1,2,3,4" for scoring or just a single intiger
    # Split the string by commas and convert each part to an integer
    if (numbers == ""):
        return 0
    if isinstance(numbers, str):
        numbers = [int(num.strip()) for num in numbers.split(',')]
        # Return the highest number
        return max(numbers)
    return numbers

def parseLeave(leave):
    if leave == "Yes":
        return True
    if leave == "No":
        return False
    
def getStringSeparated(unsortedString: str):
    sortedStrings = []
    baseOfSubString = 0
    for i, char in enumerate(unsortedString):
        newString = ""
        if((char == ',') or (i == unsortedString.__len__() - 1)):
            for e in range(i-baseOfSubString + 1):
                if((not (unsortedString[e+baseOfSubString] == ',')) and (not (unsortedString[e+baseOfSubString] == ' '))):
                    newString = newString+unsortedString[e+baseOfSubString]
            sortedStrings.append(newString)
            baseOfSubString = i+1
    return sortedStrings

quantative_values = {
    "Bad": 0,
    "Ok": 0.5,
    "Good": 1
}
    


with open(input_file_name, "r", newline="") as input_csv_file:

    # Get rid of any existing output file
    if os.path.exists(output_file_name):
        os.remove(output_file_name)
    
    input_handling_object = csv.reader(input_csv_file)

    for row_num, row_data in enumerate(input_handling_object):
    # Skip the 0th row (column titles). Python starts counting at 0 instead of 1
        if row_num > -1:
            # Put all of the information from a single row in excel into a python object
            # called "team_match_entry" to make it easier to deal with later on
            # print(f"Processing Row Number {row_num + 1}")
            # strings = getStringSeparated(row_data[8])
            # print(strings[2])
            strings = getStringSeparated(row_data[8])

            defendedScoringBool = False
            defendedIntakingBool = False
            defendedPathingBool = False

            for string in strings:  
                match string:
                    case "Blockotherteamsfromscoring":
                        defendedScoringBool = True
                    case "Blockotherteamsfrompickingupfuel":
                        defendedIntakingBool = True
                    case "Blockthepathofotherrobotscrossingthefield":
                        defendedPathingBool = True

            team_match_entry = shared_classes.SingleTeamSingleMatchEntry(
                commenter = row_data[2],
                team_num = parse_team_number(row_data[3]),
                qual_match_num=parse_match_number(row_data[4]),
                autoFuel = parse_match_number(row_data[5]),
                autoL1Climb = True if row_data[6] == "Yes" else False,

                teleFuel = parse_match_number(row_data[7]),

                defenseOnScoring = defendedScoringBool,
                defenseOnIntaking = defendedIntakingBool,
                defenseOnPathing = defendedPathingBool,

                l1Climb = True if row_data[9] == "L1" else False,
                l2Climb = True if row_data[9] == "L2" else False,
                l3Climb = True if row_data[9] == "L3" else False,

                auto = quantative_values.get(row_data[10]),
                speed = quantative_values.get(row_data[11]),
                passes = quantative_values.get(row_data[12]),
                pickupSpeed = quantative_values.get(row_data[13]),
                scoringSpeed = quantative_values.get(row_data[14]),
                driverDecisiveness = quantative_values.get(row_data[15]),
                balance = quantative_values.get(row_data[16]),
                wouldYouPick = quantative_values.get(row_data[17]),

                robotBroke = True if row_data[18] == "Yes - if yes please add details in your comments" else False,
                comment = row_data[19]
            )
            # Add the single-team single-match entry (i.e. data from one row) to a list
            # containing all of these entries
            all_team_match_entries.append(team_match_entry)
        else:
            print("Skipping Row Number 1 (Column Titles)")

    # Go through every match entry one-by-one and check if a class for all of a team's matches
# has been created yet (the TeamData class). If not, create it and add it to a list of
# these classes. Then, add the current match entry to its corresponding class containing all
# of that team's match entries.

for match_entry in all_team_match_entries:
    # If the team num is -1 (due to an input error), skip this iteration of the for loop
    if match_entry.team_num == -1:
        continue

    if match_entry.team_num not in team_num_list:
        team_num_list.append(match_entry.team_num)

        new_single_team_data = shared_classes.TeamData()
        new_single_team_data.team_num = match_entry.team_num
        # print(new_single_team_data.team_num)

        all_team_data[new_single_team_data.team_num] = new_single_team_data

    for single_teams_data in all_team_data.values():
        if single_teams_data.team_num == match_entry.team_num:

            duplicate_entry = False

            for single_team_match in single_teams_data.match_data:
                if match_entry.qual_match_num == single_team_match.qual_match_num:
                    duplicate_entry = True

            if not duplicate_entry:
                single_teams_data.match_data.append(match_entry)



with xlsxwriter.Workbook(output_file_name) as output_workbook:
    percent_format = output_workbook.add_format({'num_format': '0.0%'})
    one_decimal_format = output_workbook.add_format({'num_format': '0.0'})

    team_num_list = sorted(team_num_list)

    # Create the ranking worksheet first
    ranking_worksheet = output_workbook.add_worksheet("Rankings")
    accuracy_worksheet = output_workbook.add_worksheet("Accuracy")
    output_worksheets.append(ranking_worksheet)
    # output_worksheets.append(accuracy_worksheet)

    # Create a new sheet for every team
    for team_num in team_num_list:
        points = []
        auto_points = []
        tele_points = []
        matches = []
        single_teams_worksheet = output_workbook.add_worksheet(str(team_num))
        output_worksheets.append(single_teams_worksheet)

        # Find the all_team_data entry for the team with the same number as the team_num
        # variable
        for single_teams_data in all_team_data.values():
            if single_teams_data.team_num == team_num:

                points.clear
                auto_points.clear
                tele_points.clear
                matches.clear
                single_teams_data.match_data = sorted(
                single_teams_data.match_data,
                key=lambda x: x.qual_match_num,
                reverse=False)
                # first number is height and second is left and right
                single_teams_worksheet.write(0, 2, "Ranks")
                single_teams_worksheet.write(1, 0, "Total Points")
                single_teams_worksheet.write(2, 0, "Tele Points")
                single_teams_worksheet.write(3, 0, "Auto Points")
                single_teams_worksheet.write(4, 0, "Coral Points")
                single_teams_worksheet.write(5, 0, "Algae Points")
                single_teams_worksheet.write(6, 0, "Rice Score")
                single_teams_worksheet.write(7, 0, "--------------------------------------------------------------------------------------------------")
                single_teams_worksheet.write(8, 0, "Qualitative")
                single_teams_worksheet.write(0, 4, "Broke?")
                single_teams_worksheet.write(0, 5, "Match #")
                single_teams_worksheet.write(0, 6, "Scouter")
                single_teams_worksheet.write(0, 7, "Comment")
                single_teams_worksheet.write(38, 0, "Match")
                single_teams_worksheet.write(38, 1, "Auto")
                single_teams_worksheet.write(38, 2, "Tele")
                single_teams_worksheet.write(38, 3, "Person")
                with open("input2.csv.", "r", newline="", encoding="utf8") as input_csv_file:
                    input_handling_object = csv.reader(input_csv_file)
                    for row_num, row_data in enumerate(input_handling_object):
                        # row_num = 0
                        # print(team_num)
                        # print(row_data[2])
                        if not parse_team_number(row_data[2]) == team_num:
                            # row_num = row_num + 1
                            1==1
                        else:
                            single_teams_worksheet.write(9, 0, row_data[5])
                            single_teams_data.drivetrain = row_data[5]
                            if row_data[5] == "Swerve":
                                single_teams_data.swerve = True
                            else:
                                single_teams_data.swerve = False
                            single_teams_data.fuelCapacity = row_data[7]

                for i, match in enumerate(single_teams_data.match_data):

                    # for ever match the team played add the point now and avereges the by times itterated after
                    single_teams_data.aveAutoClimbPoints += 15 if match.autoL1Climb else 0
                    
                    single_teams_data.aveAutoFuelPoints += match.autoFuel

                    single_teams_data.aveTeleFuelPoint += match.teleFuel

                    single_teams_data.aveTeleClimbPoints += 10 if match.l1Climb else 0
                    single_teams_data.aveTeleClimbPoints += 20 if match.l2Climb else 0
                    single_teams_data.aveTeleClimbPoints += 30 if match.l3Climb else 0

                    single_teams_data.robotBroke.append(match.robotBroke)
                    single_teams_data.comments.append(match.comment)
                    single_teams_data.commenters.append(match.commenter)


                    single_teams_data.avePasses += match.passes
                    single_teams_data.aveAuto += match.auto
                    single_teams_data.aveSpeed += match.speed
                    single_teams_data.avePickupSpeed += match.pickupSpeed
                    single_teams_data.aveScoringSpeed += match.scoringSpeed
                    single_teams_data.avePickupSpeed += match.pickupSpeed
                    single_teams_data.aveDriverDecisiveness += match.driverDecisiveness
                    single_teams_data.aveBalance += match.balance
                    single_teams_data.aveWouldYouPick += match.wouldYouPick


                    matches.append(match.qual_match_num)

                    global timesIterated
                    timesIterated = i + 1
                single_teams_data.aveAutoClimbPoints = single_teams_data.aveAutoClimbPoints/timesIterated
                single_teams_data.aveAutoFuelPoints = single_teams_data.aveAutoFuelPoints/timesIterated
                single_teams_data.aveTeleFuelPoint = single_teams_data.aveTeleFuelPoint/timesIterated
                single_teams_data.aveTeleClimbPoints = single_teams_data.aveTeleClimbPoints/timesIterated

                single_teams_data.avePasses = match.passes/timesIterated
                single_teams_data.aveAuto = match.auto/timesIterated
                single_teams_data.aveSpeed = match.speed/timesIterated
                single_teams_data.avePickupSpeed = match.pickupSpeed/timesIterated
                single_teams_data.aveScoringSpeed = match.scoringSpeed/timesIterated
                single_teams_data.avePickupSpeed = match.pickupSpeed/timesIterated
                single_teams_data.aveDriverDecisiveness = match.driverDecisiveness/timesIterated
                single_teams_data.aveBalance = match.balance/timesIterated
                single_teams_data.aveWouldYouPick = match.wouldYouPick/timesIterated

                single_teams_data.aveAutoPoints = single_teams_data.aveAutoClimbPoints + single_teams_data.aveAutoFuelPoints
                single_teams_data.aveTelePoints = single_teams_data.aveTeleClimbPoints + single_teams_data.aveTeleClimbPoints
                single_teams_data.avePoints = single_teams_data.aveTelePoints + single_teams_data.aveAutoPoints

# this is all chat gpt to convert to json

# 1. Define the custom JSON encoder
class DataclassEncoder(json.JSONEncoder):
    def default(self, obj):
        if is_dataclass(obj):
            # Convert the dataclass instance to a dictionary
            return asdict(obj)
        # Let the base class default method handle other types
        return super().default(obj)


# 3. Write the list to a JSON file using the custom encoder
file_path = "outputWebsiteData.json"
try:
    with open(file_path, 'w', encoding='utf-8') as json_file:
        # Use the custom encoder with json.dump()
        json.dump(all_team_data, json_file, cls=DataclassEncoder, indent=4)
    print(f"Successfully wrote data to {file_path}")
except IOError as e:
    print(f"An error occurred while writing the file: {e}")

                # single_teams_data.riceScore = (single_teams_data.aveAutoPoints*0.333)+(single_teams_data.aveBargePoints*0.333)+(((single_teams_data.aveSpeed)
                # +(single_teams_data.aveDriver) * 5) / 0.333)

                # single_teams_worksheet.write(8, 2, single_teams_data.quantativeAve)

                # if (single_teams_data.aveCoralPoints > 10):
                #     single_teams_data.coral = "✅"
                # else:
                #     single_teams_data.coral = "❌"
                
                # if (single_teams_data.aveAlgaePoints > 2):
                #     single_teams_data.algae = "✅"
                # else:
                #     single_teams_data.algae = "❌"
                # if (single_teams_data.aveBargePoints > 5):
                #     single_teams_data.climb = "✅"
                # else:
                #     single_teams_data.climb = "❌"


                # allClimbs = (single_teams_data.deepClimb + single_teams_data.shallowClimb +
                # single_teams_data.park + single_teams_data.noClimb)
                # climb_categories = ["Climb on the deep cage", "Climb on the shallow cage", "Park", "No Climb", ]
                # climb_values = [(single_teams_data.deepClimb/allClimbs),(single_teams_data.shallowClimb/allClimbs),
                # (single_teams_data.park/allClimbs),(single_teams_data.noClimb/allClimbs)]

                
                # for i, (category, value) in enumerate(zip(climb_categories, climb_values)):
                #     single_teams_worksheet.write(DATA_START_ROW + 2 + i, 19, category)
                #     single_teams_worksheet.write(DATA_START_ROW + 2 + i, 20, value)     

                # # Create the pie chart
                # barge_pie_chart = output_workbook.add_chart({'type': 'pie'})
                # barge_pie_chart.set_title({'name': 'Endgame Climb Distribution'})
                # barge_pie_chart.add_series({
                #     'name': 'Climb Level',
                #     'categories': f'={single_teams_data.team_num}!T{DATA_START_ROW + 2}:T{DATA_START_ROW + 6}',
                #     'values': f'={single_teams_data.team_num}!U{DATA_START_ROW + 2}:U{DATA_START_ROW + 6}',
                #     'data_labels': {'percentage': True},
                #     'points': [
                #         {'fill': {'color': chart_colors["RED"]}},
                #         {'fill': {'color': chart_colors["YELLOW"]}},
                #         {'fill': {'color': chart_colors["BLUE"]}},
                #         {'fill': {'color': chart_colors["GREEN"]}},
                #     ]
                # })

                # single_teams_worksheet.insert_chart(f"{FIRST_CHART_COL}{CHART_START_ROW + CHART_ROW_SPACING}", barge_pie_chart)

                # # Write the match numbers and points to the worksheet

                # max_matches = 94
                # limit_matches = matches[:max_matches]
                # limit_auto_points = auto_points[:max_matches]
                # limit_tele_points = tele_points[:max_matches]
                
                # for i, aPoints in enumerate(limit_auto_points):
                #     # print(f"Writing {aPoints} to row {DATA_START_ROW + 1 + i}")
                #     single_teams_worksheet.write(DATA_START_ROW + 2 + i, 1, aPoints)

                # for i, match_num in enumerate(limit_matches):
                #     if (match_num < 2):
                #         1==1
                #     else:
                #         single_teams_worksheet.write(DATA_START_ROW + 2 + i, 0, match_num)
                
                # for i, tPoints in enumerate(limit_tele_points):
                #     # print(f"Writing {tPoints} to row {DATA_START_ROW + 1 + i}")
                #     single_teams_worksheet.write(DATA_START_ROW + 2 + i, 2, tPoints)

                # # Create the line chart
                # a_points_line_chart = output_workbook.add_chart({'type': 'line'})
                # a_points_line_chart.set_title({'name': 'Auto Points Over Time'})
                # a_points_line_chart.set_x_axis({'name': 'Match Number'})
                # a_points_line_chart.set_y_axis({'name': 'Auto Points'})
                
                # num_matches = matches[len(matches) -1]

                # a_points_line_chart.add_series({
                #     'name': 'Points',
                #     'categories': f'={team_num}!$A${DATA_START_ROW + 2}:$A${DATA_START_ROW + 1 + num_matches}',
                #     'values': f'={team_num}!$B${DATA_START_ROW + 2}:$B${DATA_START_ROW + 1 + num_matches}',
                #     'line': {'color': 'blue'}
                # })

                # line_chart_row = CHART_START_ROW + CHART_ROW_SPACING  # Adjust the row number as needed
                # single_teams_worksheet.insert_chart(f"{THIRD_CHART_COL}{CHART_START_ROW + CHART_ROW_SPACING}", a_points_line_chart)


                # # Create the line chart
                # t_points_line_chart = output_workbook.add_chart({'type': 'line'})
                # t_points_line_chart.set_title({'name': 'Tele Points Over Time'})
                # t_points_line_chart.set_x_axis({'name': 'Match Number'})
                # t_points_line_chart.set_y_axis({'name': 'Tele Points'})
                
                # num_matches = matches[len(matches) - 1]

                # t_points_line_chart.add_series({
                #     'name': 'Points',
                #     'categories': f'={team_num}!$A${DATA_START_ROW + 2}:$A${DATA_START_ROW + 1 + num_matches}',
                #     'values': f'={team_num}!$C${DATA_START_ROW + 2}:$C${DATA_START_ROW + 1 + num_matches}',
                #     'line': {'color': 'blue'}
                # })

                # line_chart_row = CHART_START_ROW + CHART_ROW_SPACING  # Adjust the row number as needed
                # single_teams_worksheet.insert_chart(f"{"Q"}{line_chart_row }", t_points_line_chart)


    # does all the rankings
    # for i, team in enumerate(all_team_data.values()):
    #     inserted = "no"
    #     # print(team.team_num)
    #     # print(team.aveTelePoints)
    #     for rank in range(len(total_points_rankings)):
    #         if total_points_rankings[rank] < team.avePoints:
    #             if inserted == "no":
    #                 total_points_rankings.insert(rank, team.avePoints)
    #                 total_points_rankings_team_names.insert(rank, team.team_num)
    #                 inserted = "yes"
    #             # print(total_points_rankings[rank])
    #     if inserted == "no":
    #         total_points_rankings.append(team.avePoints)
    #         total_points_rankings_team_names.append(team.team_num)

    #     inserted = "no"
    #     for rank in range(len(auto_score_rankings)):
    #         if auto_score_rankings[rank] < team.aveAutoPoints:
    #             if inserted == "no":
    #                 auto_score_rankings.insert(rank, team.aveAutoPoints)
    #                 auto_score_rankings_team_names.insert(rank, team.team_num)
    #                 inserted = "yes"
    #     if inserted == "no":
    #         auto_score_rankings.append(team.aveAutoPoints)
    #         auto_score_rankings_team_names.append(team.team_num)

    #     inserted = "no"
    #     for rank in range(len(tele_score_rankings)):
    #         if tele_score_rankings[rank] < team.aveTelePoints:
    #             if inserted == "no":
    #                 tele_score_rankings.insert(rank, team.aveTelePoints)
    #                 tele_score_rankings_team_names.insert(rank, team.team_num)
    #                 inserted = "yes"
    #     if inserted == "no":
    #         tele_score_rankings.append(team.aveTelePoints)
    #         tele_score_rankings_team_names.append(team.team_num)

    #     inserted = "no"
    #     for rank in range(len(rice_score_rankings)):
    #         if rice_score_rankings[rank] < team.riceScore:
    #             if inserted == "no":
    #                 rice_score_rankings.insert(rank, team.riceScore)
    #                 rice_scores.insert(rank, team.riceScore)
    #                 inserted = "yes"
    #     if inserted == "no":
    #         rice_score_rankings.append(team.riceScore)
    #         rice_score_rankings_team_names.append(team.team_num)
    #         rice_scores.append(team.riceScore)

    # ranking_worksheet.write(0, 0, "S,C,A =")
    # ranking_worksheet.write(1, 0, "Swerve")
    # ranking_worksheet.write(2, 0, "Coral")
    # ranking_worksheet.write(3, 0, "Algae")
    # ranking_worksheet.write(0, 2, "Total")
    # ranking_worksheet.write(0, 3, "S,C,A")
    # ranking_worksheet.write(0, 4, "Auto")
    # ranking_worksheet.write(0, 5, "S,C,A")
    # ranking_worksheet.write(0, 6, "Tele")
    # ranking_worksheet.write(0, 7, "S,C,A")
    # ranking_worksheet.write(0, 8, "Coral")
    # ranking_worksheet.write(0, 9, "S,C,A")
    # ranking_worksheet.write(0, 10, "Algae")
    # ranking_worksheet.write(0, 11, "S,C,A")
    # ranking_worksheet.write(0, 12, "Rice Score")
    # ranking_worksheet.write(0, 13, "S,C,A, CLIMB")
    # ranking_worksheet.write(0, 14, "Rice Score")
    # if len(all_team_data.values()) < 75:
    #     for i in range(len(all_team_data.values())):
    #         ranking_worksheet.write(i + 1, 2, total_points_rankings_team_names[i])
    #         ranking_worksheet.write(i + 1, 4, auto_score_rankings_team_names[i])
    #         ranking_worksheet.write(i + 1, 6, tele_score_rankings_team_names[i])
    #         ranking_worksheet.write(i + 1, 12, rice_score_rankings_team_names[i])
    #         ranking_worksheet.write(i + 1, 15, rice_scores[i])
    # for i, worksheet in enumerate(output_worksheets):
    #     if not i == 0:
    #         for e in range(len(total_points_rankings_team_names)):
    #             if team_num_list[i-1] == total_points_rankings_team_names[e]:
    #                 worksheet.write(1, 2, e+1)
    #         for e in range(len(team_num_list)):
    #             if team_num_list[i-1] ==auto_score_rankings_team_names[e]:
    #                 worksheet.write(2, 2, e+1)
    #         for e in range(len(team_num_list)):
    #             if team_num_list[i-1] == tele_score_rankings_team_names[e]:
    #                 worksheet.write(3, 2, e+1)
    #         for e in range(len(coral_score_rankings_team_names)):
    #             if team_num_list[i-1] == coral_score_rankings_team_names[e]:
    #                 worksheet.write(4, 2, e+1)
    #         for e in range(len(team_num_list)):
    #             if team_num_list[i-1] == algae_score_rankings_team_names[e]:
    #                 worksheet.write(5, 2, e+1)
    #         for e in range(len(team_num_list)):
    #             if team_num_list[i-1] == rice_score_rankings_team_names[e]:
    #                 worksheet.write(6, 2, e+1)
            
            # for e, team in enumerate(team_num_list):
            #     if team == total_points_rankings_team_names[e]:
            #         worksheet.write(1, 3, total_points_rankings[e])

                # first number is height and second is left and right
                # single_teams_worksheet.write(0, 2, "Ranks")
                # single_teams_worksheet.write(1, 0, "Total Points")
                # single_teams_worksheet.write(2, 0, "Auto Points")
                # single_teams_worksheet.write(3, 0, "Tele Points")
                # single_teams_worksheet.write(4, 0, "Qualitative")
                # single_teams_worksheet.write(0, 5, "Match #")
                # single_teams_worksheet.write(0, 6, "Scouter")
                # single_teams_worksheet.write(0, 7, "Comment")
    # tbaSorting = tba_match_sorting
    # green_format = output_workbook.add_format()
    # green_format.set_pattern(1)
    # green_format.set_bg_color('green')

    # yellow_format = output_workbook.add_format()
    # yellow_format.set_pattern(1)
    # yellow_format.set_bg_color('yellow')

    # red_format = output_workbook.add_format()
    # red_format.set_pattern(1)
    # red_format.set_bg_color('red')

    # purple_format = output_workbook.add_format()
    # purple_format.set_pattern(1)
    # purple_format.set_bg_color('purple')

    # white_format = output_workbook.add_format()
    # white_format.set_pattern(1)
    # white_format.set_bg_color('white')

    # accuracy_worksheet.write(0, 1, "Match")
    # accuracy_worksheet.write(0, 2, "Color")
    # accuracy_worksheet.write(0, 3, "Overall%")
    # accuracy_worksheet.write(0, 4, "Auto%")
    # accuracy_worksheet.write(0, 5, "Tele%")
    # accuracy_worksheet.write(0, 6, "Climb%")
    # accuracy_worksheet.write(0, 8, "robot1%")
    # accuracy_worksheet.write(0, 9, "robot2%")
    # accuracy_worksheet.write(0, 10, "robot3%")
    # accuracy_worksheet.write(0, 11, "scout1%")
    # accuracy_worksheet.write(0, 12, "scout2%")
    # accuracy_worksheet.write(0, 13, "scout3%")
    # tba_match_sorting.initializeTBAData()
    # def outputFormatWithTolerance(number, yellowTolerance, redTolerance):
    #     format = green_format
    #     if number > 40:
    #         if number > 70:
    #             format = red_format
    #         else:
    #             format = yellow_format
    #     return format
    # def outputFormatMissedClimb(missedClimb):
    #     if missedClimb:
    #         return purple_format
    #     else:
    #         return white_format
    # for matchNum in range(max_matches):
    #     # teamsInAMatch = [all_team_match_entries[0],all_team_match_entries[1],all_team_match_entries[2],all_team_match_entries[3],all_team_match_entries[4],all_team_match_entries[5]]
    #     teamsInAMatch = []
    #     for match_entry in all_team_match_entries:
    #         # print(match_entry.qual_match_num)
    #         # print(matchNum)
    #         if match_entry.qual_match_num == matchNum+1:
    #             # print(str(match_entry.qual_match_num)+ " and " + str(matchNum+1))
    #             teamsInAMatch.append(match_entry)
    #     # print(len(teamsInAMatch))
    #     allMatchesValidationData.append(tba_match_sorting.makeBothAllianceMatchClass(teamsInAMatch,all_team_data))
    #     redRow = ((matchNum+1) * 2) - 1
    #     blueRow = (matchNum+1) * 2
    #     accuracy_worksheet.write(redRow, 1, allMatchesValidationData[matchNum].matchNumRed)
    #     accuracy_worksheet.write(redRow, 2, "Red")

    #     redInacuracy = allMatchesValidationData[matchNum].overallInaccuracyRed
    #     accuracy_worksheet.write(redRow, 3, redInacuracy, outputFormatWithTolerance(redInacuracy,40,70))
    #     accuracy_worksheet.write(redRow, 4, allMatchesValidationData[matchNum].autoInaccuracyRed)
    #     accuracy_worksheet.write(redRow, 5, allMatchesValidationData[matchNum].teleInaccuracyRed)
    #     accuracy_worksheet.write(redRow, 6, allMatchesValidationData[matchNum].endGameInaccuracyRed)

    #     scouter1In = allMatchesValidationData[matchNum].scouterOneInacuracyRed
    #     scouter2In = allMatchesValidationData[matchNum].scouterTwoInacuracyRed
    #     scouter3In = allMatchesValidationData[matchNum].scouterThreeInacuracyRed

    #     accuracy_worksheet.write(redRow, 8, scouter1In, outputFormatWithTolerance(scouter1In,40,70))
    #     accuracy_worksheet.write(redRow, 9, scouter2In, outputFormatWithTolerance(scouter2In,40,70))
    #     accuracy_worksheet.write(redRow, 10, scouter3In, outputFormatWithTolerance(scouter3In,40,70))
    #     accuracy_worksheet.write(redRow, 11, allMatchesValidationData[matchNum].scouterOneNameRed,outputFormatMissedClimb(allMatchesValidationData[matchNum].scouterOneMissClimbRed))
    #     accuracy_worksheet.write(redRow, 12, allMatchesValidationData[matchNum].scouterTwoNameRed,outputFormatMissedClimb(allMatchesValidationData[matchNum].scouterTwoMissClimbRed))
    #     accuracy_worksheet.write(redRow, 13, allMatchesValidationData[matchNum].scouterThreeNameRed,outputFormatMissedClimb(allMatchesValidationData[matchNum].scouterThreeMissClimbRed))

    #     accuracy_worksheet.write(blueRow, 1, allMatchesValidationData[matchNum].matchNumBlue)
    #     accuracy_worksheet.write(blueRow, 2, "Blue")
    #     blueInacuracy = allMatchesValidationData[matchNum].overallInaccuracyBlue
    #     accuracy_worksheet.write(blueRow, 3, blueInacuracy, outputFormatWithTolerance(blueInacuracy,40,70))
    #     accuracy_worksheet.write(blueRow, 4, allMatchesValidationData[matchNum].autoInaccuracyBlue)
    #     accuracy_worksheet.write(blueRow, 5, allMatchesValidationData[matchNum].teleInaccuracyBlue)
    #     accuracy_worksheet.write(blueRow, 6, allMatchesValidationData[matchNum].endGameInaccuracyBlue)

    #     scouter1In = allMatchesValidationData[matchNum].scouterOneInacuracyBlue
    #     scouter2In = allMatchesValidationData[matchNum].scouterTwoInacuracyBlue
    #     scouter3In = allMatchesValidationData[matchNum].scouterThreeInacuracyBlue
    #     accuracy_worksheet.write(blueRow, 8, scouter1In, outputFormatWithTolerance(scouter1In,40,70))
    #     accuracy_worksheet.write(blueRow, 9, scouter2In, outputFormatWithTolerance(scouter2In,40,70))
    #     accuracy_worksheet.write(blueRow, 10, scouter3In, outputFormatWithTolerance(scouter3In,40,70))
    #     accuracy_worksheet.write(blueRow, 11, allMatchesValidationData[matchNum].scouterOneNameBlue,outputFormatMissedClimb(allMatchesValidationData[matchNum].scouterOneMissClimbBlue))
    #     accuracy_worksheet.write(blueRow, 12, allMatchesValidationData[matchNum].scouterTwoNameBlue,outputFormatMissedClimb(allMatchesValidationData[matchNum].scouterTwoMissClimbBlue))
    #     accuracy_worksheet.write(blueRow, 13, allMatchesValidationData[matchNum].scouterThreeNameBlue,outputFormatMissedClimb(allMatchesValidationData[matchNum].scouterThreeMissClimbBlue))
    # print("Outputsheet done!!!")