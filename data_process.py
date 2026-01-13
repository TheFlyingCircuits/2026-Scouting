import csv
import os
import xlsxwriter
import tba_match_sorting
import shared_classes

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
coral_score_rankings = []
algae_score_rankings = []
rice_score_rankings = []
total_points_rankings_team_names = []
tele_score_rankings_team_names = []
auto_score_rankings_team_names = []
coral_score_rankings_team_names = []
algae_score_rankings_team_names = []
rice_score_rankings_team_names = []
total_points_rankings_SCA = []
tele_score_rankings_SCA = []
auto_score_rankings_SCA = []
coral_score_rankings_SCA = []
algae_score_rankings_SCA = []
rice_score_rankings_SCA = []
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

leavePointsValue = 3
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
            team_match_entry = shared_classes.SingleTeamSingleMatchEntry(
                commenter = row_data[2],
                team_num = parse_team_number(row_data[3]),
                qual_match_num=parse_match_number(row_data[4]),
                leave = parseLeave(row_data[5]),
                autoL4 = get_highest_number(row_data[6]),
                autoL3 = get_highest_number(row_data[7]),
                autoL2 = get_highest_number(row_data[8]),
                autoL1 = get_highest_number(row_data[9]),
                autoProcessor = get_highest_number(row_data[10]),
                autoNet = get_highest_number(row_data[11]),
                algaeRemoved = get_highest_number(row_data[12] + row_data[19]),
                teleL4 = get_highest_number(row_data[13]),
                teleL3 = get_highest_number(row_data[14]),
                teleL2 = get_highest_number(row_data[15]),
                teleL1 = get_highest_number(row_data[16]),
                teleProcessor = get_highest_number(row_data[17]),
                teleNet = get_highest_number(row_data[18]),
                climb = row_data[20],

                auto = row_data[24],
                speed = row_data[25],
                pickupSpeed = row_data[26],
                scoring = row_data[27],
                driverDecisiveness = row_data[28],
                balance = row_data[29],
                wouldYouPick = row_data[30],
                robotBroke = row_data[31],
                comment = row_data[32],

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
                                single_teams_data.swerve = "✅"
                            else:
                                single_teams_data.swerve = "❌"

                            # single_teams_worksheet.write(5, 2, row_data[5])
                            # single_teams_worksheet.write(6, 2, row_data[7])
                            # single_teams_worksheet.write(7, 2, row_data[17])
                            # single_teams_worksheet.write(8, 2, row_data[18])
                            # single_teams_worksheet.write(9, 2, row_data[19])
                            # single_teams_worksheet.write(10, 2, row_data[6])
                            # single_teams_worksheet.write(11, 2, row_data[20])
                            1==1

                for i, match in enumerate(single_teams_data.match_data):
                    # qualative = [single_teams_data.auto, single_teams_data.speed, single_teams_data.pickupSpeed, single_teams_data.scoring, 
                    # single_teams_data.driverDecisiveness, single_teams_data.balance, single_teams_data.wouldYouPick]
                    quantative_values = {
                        "Bad": 0,
                        "Ok": 0.5,
                        "Good": 1
                    }

                    auto = quantative_values.get(match.auto)
                    speed = quantative_values.get(match.speed)
                    pickupSpeed = quantative_values.get(match.pickupSpeed)
                    scoring = quantative_values.get(match.scoring)
                    driverDecisiveness = quantative_values.get(match.driverDecisiveness)
                    balance = quantative_values.get(match.balance)
                    wouldYouPick = quantative_values.get(match.wouldYouPick)

                    # print(auto)
                    # print(speed)
                    # print(pickupSpeed)
                    single_teams_data.quantativeAve = single_teams_data.quantativeAve + (
                        auto + speed + pickupSpeed +
                        scoring + driverDecisiveness + balance + wouldYouPick) / 7

                    if match.robotBroke == "No":
                        single_teams_worksheet.write(single_teams_data.commentNum, 4, "❌")
                    else:
                        single_teams_worksheet.write(single_teams_data.commentNum, 4, "✅")
                    single_teams_worksheet.write(single_teams_data.commentNum, 5, match.qual_match_num)
                    single_teams_worksheet.write(single_teams_data.commentNum, 6, match.commenter)
                    single_teams_worksheet.write(single_teams_data.commentNum, 7, match.comment)
                    single_teams_worksheet.write(single_teams_data.commentNum + 38, 3, match.commenter)
                    single_teams_data.commentNum =  single_teams_data.commentNum + 1
                        # for ever match the team played add the point now and avereges the by times itterated after
                    if match.leave == True:
                        # print(match.leave)
                        single_teams_data.aveLeavePoints = single_teams_data.aveLeavePoints + leavePointsValue
                    single_teams_data.aveSpeed = speed
                    single_teams_data.aveDriver = driverDecisiveness
                    single_teams_data.aveAutoL4Points = single_teams_data.aveAutoL4Points + (match.autoL4 * autoL4PointsValue)
                    single_teams_data.aveAutoL3Points = single_teams_data.aveAutoL3Points + (match.autoL3 * autoL3PointsValue)
                    single_teams_data.aveAutoL2Points = single_teams_data.aveAutoL2Points + (match.autoL2 * autoL2PointsValue)
                    single_teams_data.aveAutoL1Points = single_teams_data.aveAutoL1Points + (match.autoL1 * autoL1PointsValue)
                    single_teams_data.aveAutoProcessorPoints = single_teams_data.aveAutoProcessorPoints + (match.autoProcessor * 
                    processorPointsValue)
                    single_teams_data.aveAutoNetPoints = single_teams_data.aveAutoNetPoints + (match.autoNet * netPointsValue)
                    single_teams_data.aveTeleL4Points = single_teams_data.aveTeleL4Points + (match.teleL4 * L4PointsValue)
                    single_teams_data.aveTeleL3Points = single_teams_data.aveTeleL3Points + (match.teleL3 * L3PointsValue)
                    single_teams_data.aveTeleL2Points = single_teams_data.aveTeleL2Points + (match.teleL2 * L2PointsValue)
                    single_teams_data.aveTeleL1Points = single_teams_data.aveTeleL1Points + (match.teleL1 * L1PointsValue)
                    single_teams_data.aveTeleProcessorPoints = single_teams_data.aveTeleProcessorPoints + (match.teleProcessor * 
                    processorPointsValue)
                    single_teams_data.aveTeleNetPoints = single_teams_data.aveTeleNetPoints + (match.teleNet * netPointsValue)
                    climb: int = 0
                    if match.climb == "Climb on the deep cage":
                        single_teams_data.aveBargePoints = single_teams_data.aveBargePoints + 12
                        single_teams_data.deepClimb = single_teams_data.deepClimb + 1
                        climb = 12
                    elif match.climb == "Climb on the shallow cage":
                        single_teams_data.aveBargePoints = single_teams_data.aveBargePoints + 6
                        single_teams_data.shallowClimb = single_teams_data.shallowClimb + 1
                        climb = 6
                    elif match.climb == "Park in the barge zone":
                        single_teams_data.aveBargePoints = single_teams_data.aveBargePoints + 2
                        single_teams_data.park = single_teams_data.park + 1
                        climb = 2
                    else:
                        single_teams_data.noClimb = single_teams_data.noClimb + 1
                        climb = 0
                    
                    single_teams_data.aveAutoPoints = (single_teams_data.aveLeavePoints + 
                    single_teams_data.aveAutoL4Points + single_teams_data.aveAutoL3Points + single_teams_data.aveAutoL2Points +
                    single_teams_data.aveAutoL1Points + single_teams_data.aveAutoProcessorPoints + single_teams_data.aveAutoNetPoints)

                    single_teams_data.aveTelePoints = (single_teams_data.aveTeleL4Points + single_teams_data.aveTeleL3Points + 
                    single_teams_data.aveTeleL2Points + single_teams_data.aveTeleL1Points +
                    single_teams_data.aveTeleProcessorPoints + single_teams_data.aveTeleNetPoints + single_teams_data.aveBargePoints)

                    single_teams_data.avePoints = single_teams_data.aveAutoPoints + single_teams_data.aveTelePoints

                    single_teams_data.aveAlgaeRemoved = single_teams_data.aveAlgaeRemoved + match.algaeRemoved

                    points.append(leavePointsValue+(match.autoL4 * autoL4PointsValue)+(match.autoL3 * autoL3PointsValue)+(match.autoL2 * autoL2PointsValue) +
                    (match.autoL1 * autoL1PointsValue)+(match.autoProcessor * processorPointsValue)+(match.autoNet * netPointsValue)+
                    (match.teleL4 * L4PointsValue)+(match.teleL3 * L3PointsValue)+(match.teleL2 * L2PointsValue)+(match.teleL1 * L1PointsValue)+
                    (match.teleProcessor * processorPointsValue)+(match.teleNet * netPointsValue)+climb)
                    
                    auto_points.append((leavePointsValue+(match.autoL4 * autoL4PointsValue)+(match.autoL3 * autoL3PointsValue)+(match.autoL2 * autoL2PointsValue) +
                    (match.autoL1 * autoL1PointsValue)+(match.autoProcessor * processorPointsValue)+(match.autoNet * netPointsValue)))
                    
                    tele_points.append(((match.teleL4 * L4PointsValue)+(match.teleL3 * L3PointsValue)+(match.teleL2 * L2PointsValue)+(match.teleL1 * L1PointsValue)+
                    (match.teleProcessor * processorPointsValue)+(match.teleNet * netPointsValue)+climb))
                    
                    matches.append(match.qual_match_num)

                    global timesIterated
                    timesIterated = i + 1
                single_teams_data.quantativeAve = single_teams_data.quantativeAve/timesIterated
                single_teams_data.aveAlgaeRemoved = single_teams_data.aveAlgaeRemoved/timesIterated
                single_teams_data.aveLeavePoints = single_teams_data.aveLeavePoints/timesIterated
                single_teams_data.aveAutoL4Points = single_teams_data.aveAutoL4Points/timesIterated
                single_teams_data.aveAutoL3Points = single_teams_data.aveAutoL3Points/timesIterated
                single_teams_data.aveAutoL2Points = single_teams_data.aveAutoL2Points/timesIterated
                single_teams_data.aveAutoL1Points = single_teams_data.aveAutoL1Points/timesIterated
                single_teams_data.aveAutoProcessorPoints = single_teams_data.aveAutoProcessorPoints/timesIterated
                single_teams_data.aveAutoNetPoints = single_teams_data.aveAutoNetPoints/timesIterated
                single_teams_data.aveTeleL4Points = single_teams_data.aveTeleL4Points/timesIterated
                single_teams_data.aveTeleL3Points = single_teams_data.aveTeleL3Points/timesIterated
                single_teams_data.aveTeleL2Points = single_teams_data.aveTeleL2Points/timesIterated
                single_teams_data.aveTeleL1Points = single_teams_data.aveTeleL1Points/timesIterated
                single_teams_data.aveTeleProcessorPoints = single_teams_data.aveTeleProcessorPoints/timesIterated
                single_teams_data.aveTeleNetPoints = single_teams_data.aveTeleNetPoints/timesIterated
                single_teams_data.aveBargePoints = single_teams_data.aveBargePoints/timesIterated
                single_teams_data.aveAutoPoints = single_teams_data.aveAutoPoints/timesIterated
                single_teams_data.aveTelePoints = single_teams_data.aveTelePoints/timesIterated
                single_teams_data.avePoints = single_teams_data.avePoints/timesIterated
                
                single_teams_data.aveCoralPoints = single_teams_data.aveAutoL4Points + single_teams_data.aveAutoL3Points + single_teams_data.aveAutoL2Points + single_teams_data.aveAutoL1Points
                single_teams_data.aveTeleL4Points + single_teams_data.aveTeleL3Points + single_teams_data.aveTeleL2Points + single_teams_data.aveTeleL1Points
                
                single_teams_data.aveAlgaePoints = single_teams_data.aveAutoProcessorPoints + single_teams_data.aveAutoNetPoints + single_teams_data.aveTeleProcessorPoints 
                + single_teams_data.aveTeleNetPoints

                single_teams_data.riceScore = (single_teams_data.aveAutoPoints*0.333)+(single_teams_data.aveBargePoints*0.333)+(((single_teams_data.aveSpeed)
                +(single_teams_data.aveDriver) * 5) / 0.333)

                single_teams_worksheet.write(8, 2, single_teams_data.quantativeAve)

                if (single_teams_data.aveCoralPoints > 10):
                    single_teams_data.coral = "✅"
                else:
                    single_teams_data.coral = "❌"
                
                if (single_teams_data.aveAlgaePoints > 2):
                    single_teams_data.algae = "✅"
                else:
                    single_teams_data.algae = "❌"
                if (single_teams_data.aveBargePoints > 5):
                    single_teams_data.climb = "✅"
                else:
                    single_teams_data.climb = "❌"


                allClimbs = (single_teams_data.deepClimb + single_teams_data.shallowClimb +
                single_teams_data.park + single_teams_data.noClimb)
                climb_categories = ["Climb on the deep cage", "Climb on the shallow cage", "Park", "No Climb", ]
                climb_values = [(single_teams_data.deepClimb/allClimbs),(single_teams_data.shallowClimb/allClimbs),
                (single_teams_data.park/allClimbs),(single_teams_data.noClimb/allClimbs)]

                
                for i, (category, value) in enumerate(zip(climb_categories, climb_values)):
                    single_teams_worksheet.write(DATA_START_ROW + 2 + i, 19, category)
                    single_teams_worksheet.write(DATA_START_ROW + 2 + i, 20, value)     

                # Create the pie chart
                barge_pie_chart = output_workbook.add_chart({'type': 'pie'})
                barge_pie_chart.set_title({'name': 'Endgame Climb Distribution'})
                barge_pie_chart.add_series({
                    'name': 'Climb Level',
                    'categories': f'={single_teams_data.team_num}!T{DATA_START_ROW + 2}:T{DATA_START_ROW + 6}',
                    'values': f'={single_teams_data.team_num}!U{DATA_START_ROW + 2}:U{DATA_START_ROW + 6}',
                    'data_labels': {'percentage': True},
                    'points': [
                        {'fill': {'color': chart_colors["RED"]}},
                        {'fill': {'color': chart_colors["YELLOW"]}},
                        {'fill': {'color': chart_colors["BLUE"]}},
                        {'fill': {'color': chart_colors["GREEN"]}},
                    ]
                })

                single_teams_worksheet.insert_chart(f"{FIRST_CHART_COL}{CHART_START_ROW + CHART_ROW_SPACING}", barge_pie_chart)

                # Write the match numbers and points to the worksheet

                max_matches = 94
                limit_matches = matches[:max_matches]
                limit_auto_points = auto_points[:max_matches]
                limit_tele_points = tele_points[:max_matches]
                
                for i, aPoints in enumerate(limit_auto_points):
                    # print(f"Writing {aPoints} to row {DATA_START_ROW + 1 + i}")
                    single_teams_worksheet.write(DATA_START_ROW + 2 + i, 1, aPoints)

                for i, match_num in enumerate(limit_matches):
                    if (match_num < 2):
                        1==1
                    else:
                        single_teams_worksheet.write(DATA_START_ROW + 2 + i, 0, match_num)
                
                for i, tPoints in enumerate(limit_tele_points):
                    # print(f"Writing {tPoints} to row {DATA_START_ROW + 1 + i}")
                    single_teams_worksheet.write(DATA_START_ROW + 2 + i, 2, tPoints)

                # Create the line chart
                a_points_line_chart = output_workbook.add_chart({'type': 'line'})
                a_points_line_chart.set_title({'name': 'Auto Points Over Time'})
                a_points_line_chart.set_x_axis({'name': 'Match Number'})
                a_points_line_chart.set_y_axis({'name': 'Auto Points'})
                
                num_matches = matches[len(matches) -1]

                a_points_line_chart.add_series({
                    'name': 'Points',
                    'categories': f'={team_num}!$A${DATA_START_ROW + 2}:$A${DATA_START_ROW + 1 + num_matches}',
                    'values': f'={team_num}!$B${DATA_START_ROW + 2}:$B${DATA_START_ROW + 1 + num_matches}',
                    'line': {'color': 'blue'}
                })

                line_chart_row = CHART_START_ROW + CHART_ROW_SPACING  # Adjust the row number as needed
                single_teams_worksheet.insert_chart(f"{THIRD_CHART_COL}{CHART_START_ROW + CHART_ROW_SPACING}", a_points_line_chart)


                # Create the line chart
                t_points_line_chart = output_workbook.add_chart({'type': 'line'})
                t_points_line_chart.set_title({'name': 'Tele Points Over Time'})
                t_points_line_chart.set_x_axis({'name': 'Match Number'})
                t_points_line_chart.set_y_axis({'name': 'Tele Points'})
                
                num_matches = matches[len(matches) - 1]

                t_points_line_chart.add_series({
                    'name': 'Points',
                    'categories': f'={team_num}!$A${DATA_START_ROW + 2}:$A${DATA_START_ROW + 1 + num_matches}',
                    'values': f'={team_num}!$C${DATA_START_ROW + 2}:$C${DATA_START_ROW + 1 + num_matches}',
                    'line': {'color': 'blue'}
                })

                line_chart_row = CHART_START_ROW + CHART_ROW_SPACING  # Adjust the row number as needed
                single_teams_worksheet.insert_chart(f"{"Q"}{line_chart_row }", t_points_line_chart)


    # does all the rankings
    for i, team in enumerate(all_team_data.values()):
        inserted = "no"
        # print(team.team_num)
        # print(team.aveTelePoints)
        for rank in range(len(total_points_rankings)):
            if total_points_rankings[rank] < team.avePoints:
                if inserted == "no":
                    total_points_rankings.insert(rank, team.avePoints)
                    total_points_rankings_team_names.insert(rank, team.team_num)
                    total_points_rankings_SCA.insert(rank, team.swerve+team.coral+team.algae)
                    inserted = "yes"
                # print(total_points_rankings[rank])
        if inserted == "no":
            total_points_rankings.append(team.avePoints)
            total_points_rankings_team_names.append(team.team_num)
            total_points_rankings_SCA.append(team.swerve+team.coral+team.algae)

        inserted = "no"
        for rank in range(len(auto_score_rankings)):
            if auto_score_rankings[rank] < team.aveAutoPoints:
                if inserted == "no":
                    auto_score_rankings.insert(rank, team.aveAutoPoints)
                    auto_score_rankings_team_names.insert(rank, team.team_num)
                    auto_score_rankings_SCA.insert(rank, team.swerve+team.coral+team.algae)
                    inserted = "yes"
        if inserted == "no":
            auto_score_rankings.append(team.aveAutoPoints)
            auto_score_rankings_team_names.append(team.team_num)
            auto_score_rankings_SCA.append(team.swerve+team.coral+team.algae)

        inserted = "no"
        for rank in range(len(tele_score_rankings)):
            if tele_score_rankings[rank] < team.aveTelePoints:
                if inserted == "no":
                    tele_score_rankings.insert(rank, team.aveTelePoints)
                    tele_score_rankings_team_names.insert(rank, team.team_num)
                    tele_score_rankings_SCA.insert(rank, team.swerve+team.coral+team.algae)
                    inserted = "yes"
        if inserted == "no":
            tele_score_rankings.append(team.aveTelePoints)
            tele_score_rankings_team_names.append(team.team_num)
            tele_score_rankings_SCA.append(team.swerve+team.coral+team.algae)

        inserted = "no"
        for rank in range(len(coral_score_rankings)):
            if coral_score_rankings[rank] < team.aveCoralPoints:
                if inserted == "no":
                    coral_score_rankings.insert(rank, team.aveCoralPoints)
                    coral_score_rankings_team_names.insert(rank, team.team_num)
                    coral_score_rankings_SCA.insert(rank, team.swerve+team.coral+team.algae)
                    inserted = "yes"
        if inserted == "no":
            coral_score_rankings.append(team.aveCoralPoints)
            coral_score_rankings_team_names.append(team.team_num)
            coral_score_rankings_SCA.append(team.swerve+team.coral+team.algae)

        inserted = "no"
        for rank in range(len(algae_score_rankings)):
            if algae_score_rankings[rank] < team.aveAlgaePoints:
                if inserted == "no":
                    algae_score_rankings.insert(rank, team.aveAlgaePoints)
                    algae_score_rankings_team_names.insert(rank, team.team_num)
                    algae_score_rankings_SCA.insert(rank, team.swerve+team.coral+team.algae)
                    inserted = "yes"
        if inserted == "no":
            algae_score_rankings.append(team.aveAlgaePoints)
            algae_score_rankings_team_names.append(team.team_num)
            algae_score_rankings_SCA.append(team.swerve+team.coral+team.algae)

        inserted = "no"
        for rank in range(len(rice_score_rankings)):
            if rice_score_rankings[rank] < team.riceScore:
                if inserted == "no":
                    rice_score_rankings.insert(rank, team.riceScore)
                    rice_score_rankings_team_names.insert(rank, team.team_num)
                    rice_score_rankings_SCA.insert(rank, team.swerve+team.coral+team.algae+team.climb)
                    rice_scores.insert(rank, team.riceScore)
                    inserted = "yes"
        if inserted == "no":
            rice_score_rankings.append(team.riceScore)
            rice_score_rankings_team_names.append(team.team_num)
            rice_score_rankings_SCA.append(team.swerve+team.coral+team.algae+team.climb)
            rice_scores.append(team.riceScore)

    ranking_worksheet.write(0, 0, "S,C,A =")
    ranking_worksheet.write(1, 0, "Swerve")
    ranking_worksheet.write(2, 0, "Coral")
    ranking_worksheet.write(3, 0, "Algae")
    ranking_worksheet.write(0, 2, "Total")
    ranking_worksheet.write(0, 3, "S,C,A")
    ranking_worksheet.write(0, 4, "Auto")
    ranking_worksheet.write(0, 5, "S,C,A")
    ranking_worksheet.write(0, 6, "Tele")
    ranking_worksheet.write(0, 7, "S,C,A")
    ranking_worksheet.write(0, 8, "Coral")
    ranking_worksheet.write(0, 9, "S,C,A")
    ranking_worksheet.write(0, 10, "Algae")
    ranking_worksheet.write(0, 11, "S,C,A")
    ranking_worksheet.write(0, 12, "Rice Score")
    ranking_worksheet.write(0, 13, "S,C,A, CLIMB")
    ranking_worksheet.write(0, 14, "Rice Score")
    if len(all_team_data.values()) < 75:
        for i in range(len(all_team_data.values())):
            ranking_worksheet.write(i + 1, 2, total_points_rankings_team_names[i])
            ranking_worksheet.write(i + 1, 3, total_points_rankings_SCA[i])
            ranking_worksheet.write(i + 1, 4, auto_score_rankings_team_names[i])
            ranking_worksheet.write(i + 1, 5, auto_score_rankings_SCA[i])
            ranking_worksheet.write(i + 1, 6, tele_score_rankings_team_names[i])
            ranking_worksheet.write(i + 1, 7, tele_score_rankings_SCA[i])
            ranking_worksheet.write(i + 1, 8, coral_score_rankings_team_names[i])
            ranking_worksheet.write(i + 1, 9, coral_score_rankings_SCA[i])
            ranking_worksheet.write(i + 1, 10, algae_score_rankings_team_names[i])
            ranking_worksheet.write(i + 1, 11, algae_score_rankings_SCA[i])
            ranking_worksheet.write(i + 1, 12, rice_score_rankings_team_names[i])
            ranking_worksheet.write(i + 1, 13, rice_score_rankings_SCA[i])
            ranking_worksheet.write(i + 1, 15, rice_scores[i])
    for i, worksheet in enumerate(output_worksheets):
        if not i == 0:
            for e in range(len(total_points_rankings_team_names)):
                if team_num_list[i-1] == total_points_rankings_team_names[e]:
                    worksheet.write(1, 2, e+1)
            for e in range(len(team_num_list)):
                if team_num_list[i-1] ==auto_score_rankings_team_names[e]:
                    worksheet.write(2, 2, e+1)
            for e in range(len(team_num_list)):
                if team_num_list[i-1] == tele_score_rankings_team_names[e]:
                    worksheet.write(3, 2, e+1)
            for e in range(len(coral_score_rankings_team_names)):
                if team_num_list[i-1] == coral_score_rankings_team_names[e]:
                    worksheet.write(4, 2, e+1)
            for e in range(len(team_num_list)):
                if team_num_list[i-1] == algae_score_rankings_team_names[e]:
                    worksheet.write(5, 2, e+1)
            for e in range(len(team_num_list)):
                if team_num_list[i-1] == rice_score_rankings_team_names[e]:
                    worksheet.write(6, 2, e+1)
            
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
    green_format = output_workbook.add_format()
    green_format.set_pattern(1)
    green_format.set_bg_color('green')

    yellow_format = output_workbook.add_format()
    yellow_format.set_pattern(1)
    yellow_format.set_bg_color('yellow')

    red_format = output_workbook.add_format()
    red_format.set_pattern(1)
    red_format.set_bg_color('red')

    purple_format = output_workbook.add_format()
    purple_format.set_pattern(1)
    purple_format.set_bg_color('purple')

    white_format = output_workbook.add_format()
    white_format.set_pattern(1)
    white_format.set_bg_color('white')

    accuracy_worksheet.write(0, 1, "Match")
    accuracy_worksheet.write(0, 2, "Color")
    accuracy_worksheet.write(0, 3, "Overall%")
    accuracy_worksheet.write(0, 4, "Auto%")
    accuracy_worksheet.write(0, 5, "Tele%")
    accuracy_worksheet.write(0, 6, "Climb%")
    accuracy_worksheet.write(0, 8, "robot1%")
    accuracy_worksheet.write(0, 9, "robot2%")
    accuracy_worksheet.write(0, 10, "robot3%")
    accuracy_worksheet.write(0, 11, "scout1%")
    accuracy_worksheet.write(0, 12, "scout2%")
    accuracy_worksheet.write(0, 13, "scout3%")
    tba_match_sorting.initializeTBAData()
    def outputFormatWithTolerance(number, yellowTolerance, redTolerance):
        format = green_format
        if number > 40:
            if number > 70:
                format = red_format
            else:
                format = yellow_format
        return format
    def outputFormatMissedClimb(missedClimb):
        if missedClimb:
            return purple_format
        else:
            return white_format
    for matchNum in range(max_matches):
        # teamsInAMatch = [all_team_match_entries[0],all_team_match_entries[1],all_team_match_entries[2],all_team_match_entries[3],all_team_match_entries[4],all_team_match_entries[5]]
        teamsInAMatch = []
        for match_entry in all_team_match_entries:
            # print(match_entry.qual_match_num)
            # print(matchNum)
            if match_entry.qual_match_num == matchNum+1:
                # print(str(match_entry.qual_match_num)+ " and " + str(matchNum+1))
                teamsInAMatch.append(match_entry)
        # print(len(teamsInAMatch))
        allMatchesValidationData.append(tba_match_sorting.makeBothAllianceMatchClass(teamsInAMatch,all_team_data))
        redRow = ((matchNum+1) * 2) - 1
        blueRow = (matchNum+1) * 2
        accuracy_worksheet.write(redRow, 1, allMatchesValidationData[matchNum].matchNumRed)
        accuracy_worksheet.write(redRow, 2, "Red")

        redInacuracy = allMatchesValidationData[matchNum].overallInaccuracyRed
        accuracy_worksheet.write(redRow, 3, redInacuracy, outputFormatWithTolerance(redInacuracy,40,70))
        accuracy_worksheet.write(redRow, 4, allMatchesValidationData[matchNum].autoInaccuracyRed)
        accuracy_worksheet.write(redRow, 5, allMatchesValidationData[matchNum].teleInaccuracyRed)
        accuracy_worksheet.write(redRow, 6, allMatchesValidationData[matchNum].endGameInaccuracyRed)

        scouter1In = allMatchesValidationData[matchNum].scouterOneInacuracyRed
        scouter2In = allMatchesValidationData[matchNum].scouterTwoInacuracyRed
        scouter3In = allMatchesValidationData[matchNum].scouterThreeInacuracyRed

        accuracy_worksheet.write(redRow, 8, scouter1In, outputFormatWithTolerance(scouter1In,40,70))
        accuracy_worksheet.write(redRow, 9, scouter2In, outputFormatWithTolerance(scouter2In,40,70))
        accuracy_worksheet.write(redRow, 10, scouter3In, outputFormatWithTolerance(scouter3In,40,70))
        accuracy_worksheet.write(redRow, 11, allMatchesValidationData[matchNum].scouterOneNameRed,outputFormatMissedClimb(allMatchesValidationData[matchNum].scouterOneMissClimbRed))
        accuracy_worksheet.write(redRow, 12, allMatchesValidationData[matchNum].scouterTwoNameRed,outputFormatMissedClimb(allMatchesValidationData[matchNum].scouterTwoMissClimbRed))
        accuracy_worksheet.write(redRow, 13, allMatchesValidationData[matchNum].scouterThreeNameRed,outputFormatMissedClimb(allMatchesValidationData[matchNum].scouterThreeMissClimbRed))

        accuracy_worksheet.write(blueRow, 1, allMatchesValidationData[matchNum].matchNumBlue)
        accuracy_worksheet.write(blueRow, 2, "Blue")
        blueInacuracy = allMatchesValidationData[matchNum].overallInaccuracyBlue
        accuracy_worksheet.write(blueRow, 3, blueInacuracy, outputFormatWithTolerance(blueInacuracy,40,70))
        accuracy_worksheet.write(blueRow, 4, allMatchesValidationData[matchNum].autoInaccuracyBlue)
        accuracy_worksheet.write(blueRow, 5, allMatchesValidationData[matchNum].teleInaccuracyBlue)
        accuracy_worksheet.write(blueRow, 6, allMatchesValidationData[matchNum].endGameInaccuracyBlue)

        scouter1In = allMatchesValidationData[matchNum].scouterOneInacuracyBlue
        scouter2In = allMatchesValidationData[matchNum].scouterTwoInacuracyBlue
        scouter3In = allMatchesValidationData[matchNum].scouterThreeInacuracyBlue
        accuracy_worksheet.write(blueRow, 8, scouter1In, outputFormatWithTolerance(scouter1In,40,70))
        accuracy_worksheet.write(blueRow, 9, scouter2In, outputFormatWithTolerance(scouter2In,40,70))
        accuracy_worksheet.write(blueRow, 10, scouter3In, outputFormatWithTolerance(scouter3In,40,70))
        accuracy_worksheet.write(blueRow, 11, allMatchesValidationData[matchNum].scouterOneNameBlue,outputFormatMissedClimb(allMatchesValidationData[matchNum].scouterOneMissClimbBlue))
        accuracy_worksheet.write(blueRow, 12, allMatchesValidationData[matchNum].scouterTwoNameBlue,outputFormatMissedClimb(allMatchesValidationData[matchNum].scouterTwoMissClimbBlue))
        accuracy_worksheet.write(blueRow, 13, allMatchesValidationData[matchNum].scouterThreeNameBlue,outputFormatMissedClimb(allMatchesValidationData[matchNum].scouterThreeMissClimbBlue))
    print("Outputsheet done!!!")