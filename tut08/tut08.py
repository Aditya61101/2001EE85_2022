import os
import math
from datetime import datetime
start_time = datetime.now()

os.system('cls')

def get_fall(element):
    fall_at = int(element[:element.index('-')])
    return(fall_at)

def ind_innings_scorecard(ind_batter_stats, ind_score, ind_fall_of_wickets, pak_bowler_stats, ind_power_play_runs ):
    with open('Scorecard.txt','a') as scorecard:
        scorecard.write(f"\n\n{'## INDIA INNINGS': <18} {ind_score: <10}\n")
        scorecard.write(f"\n{'Batter': <23}{' ': <45}{'R': ^5}{'B': ^5}{'4s': ^5}{'6s': ^5}{'SR': >7}")
        for batter in ind_batter_stats:
            scorecard.write(f"\n{batter: <23}{ind_batter_stats[batter][-1]: <45}{ind_batter_stats[batter][0]: ^5}{ind_batter_stats[batter][1]: ^5}{ind_batter_stats[batter][2]: ^5}{ind_batter_stats[batter][3]: ^5}{ind_batter_stats[batter][4]: ^8}")

        scorecard.write('\n\nFall of Wickets\n')
        fall_statement = ''
        for outs in ind_fall_of_wickets:
            fall_statement += outs + ', '
        scorecard.write(fall_statement[:-2])

        scorecard.write(f"\n\n{'Bowler': <23}{'O': ^5}{'M': ^5}{'R': ^5}{'W': ^5}{'NB': ^5}{'WD': ^5}{'ECO': >5}")
        for bowler in pak_bowler_stats:
            scorecard.write(f"\n{bowler: <23}{pak_bowler_stats[bowler][0][-1]: ^5}{pak_bowler_stats[bowler][1]: ^5}{pak_bowler_stats[bowler][2]: ^5}{pak_bowler_stats[bowler][3]: ^5}{pak_bowler_stats[bowler][4]: ^5}{pak_bowler_stats[bowler][5]: ^5}{pak_bowler_stats[bowler][6]: ^7}")

        scorecard.write(f"\n\n{'Powerplays': <15}{'Overs': ^8}{'Runs': >8}")
        scorecard.write(f"\n{'Mandatory': <15}{'0.1-6': ^8}{ind_power_play_runs: >8}")

def ind_innings(team_pak, team_ind):
    ind_extras = 0
    ind_players = {'Batters': [], 'Bowlers': []}
    pak_players = {'Batters': [], 'Bowlers': []}
    with open('india_inns2.txt') as inns2:
        for line in inns2:
            if(line == '\n'):
                continue
            j = line.index(' ') + 1
            k = line.index(',')
            delivery = line[j:k].split('to')

            for i in range(len(delivery)):
                delivery[i] = delivery[i].strip()
            for player in team_pak:
                if(delivery[0] in player):
                    if(player not in pak_players['Bowlers']):
                        pak_players['Bowlers'].append(player)
            for player in team_ind:
                if(delivery[1] in player):
                    if(player not in ind_players['Batters']):
                        ind_players['Batters'].append(player)

        ind_batter_stats = {}
        pak_bowler_stats = {}

        for batter in ind_players['Batters']:
            ind_batter_stats[batter] = [0]*8
            ind_batter_stats[batter][-1] = 'not out'

        for bowler in pak_players['Bowlers']:
            pak_bowler_stats[bowler] = [[]]
            for i in range(6):
                pak_bowler_stats[f'{bowler}'].append(0)
        last_line = line
        ind_last_over = last_line[:last_line.index('.')+2]

    with open('india_inns2.txt') as inns2:
        i = last_line.index('.')
        last_over = float(line[:i+2]) + 1
        n_out = 0
        ind_power_play_runs = 0
        for line in inns2:
            if(line == '\n'):
                continue
            i = line.index('.')
            current_over = int(line[:i]) + 1
            current_ball = line[:i+2]
            try:
                j = line.index(' ') + 1
                k = line.index(',')
                delivery = line[j:k].split('to')
                for i in range(len(delivery)):
                    delivery[i] = delivery[i].strip()
                for player in team_pak:
                    if(delivery[0] in player):
                        current_bowler = player
                for player in team_ind:
                    if(delivery[1] in player):
                        current_batter = player
            except:
                k = 0
            if(current_over not in pak_bowler_stats[f'{current_bowler}'][0]):
                if(current_over == math.floor(last_over)):
                    last_over = float(line[:i+3]) + 1
                    pak_bowler_stats[f'{current_bowler}'][0][-1] = last_over
                else:
                    pak_bowler_stats[f'{current_bowler}'][0].append(current_over)
            try:
                try:
                    j = line.index(' ', k+1) + 1
                    k = line.index(',', k+1)
                except: 
                    k = 10000000
                try:
                    l = line.index('!')
                except:
                    l = 10000000
                run = 0
                if(k < l):
                    runs = line[j:k]
                    if(runs == 'SIX'):
                        run = 6
                        ind_batter_stats[current_batter][3] += 1
                        pak_bowler_stats[current_bowler][2] += 6
                    if(runs == 'FOUR'):
                        run = 4
                        ind_batter_stats[current_batter][2] += 1
                        pak_bowler_stats[current_bowler][2] += 4
                    if(runs == '1 run'):
                        run = 1
                        pak_bowler_stats[current_bowler][2] += 1
                    if(runs == '2 runs'):
                        run = 2
                        pak_bowler_stats[current_bowler][2] += 2
                    if(runs == '3 runs'):
                        run = 3
                        pak_bowler_stats[current_bowler][2] += 3
                    if(runs != 'wide'):
                        ind_batter_stats[current_batter][1] += 1
                    if(runs ==  'wide'): 
                        pak_bowler_stats[current_bowler][2] += 1
                        pak_bowler_stats[current_bowler][5] += 1
                        ind_extras += 1
                    if(runs ==  '2 wides'):
                        pak_bowler_stats[current_bowler][2] += 2
                        pak_bowler_stats[current_bowler][5] += 2
                        ind_extras += 2
                    if(runs ==  '3 wides'):
                        pak_bowler_stats[current_bowler][2] += 3
                        pak_bowler_stats[current_bowler][5] += 3
                        ind_extras += 3
                    if((runs == 'leg byes') | (runs == 'byes')):
                        j = line.index(' ', k+1) + 1
                        k = line.index(',', k+1)
                        runs2 = line[j:k]
                        if(runs2 == 'SIX'):
                            ind_batter_stats[current_batter][3] += 1
                            pak_bowler_stats[current_bowler][2] += 6
                            ind_extras += 6
                        if(runs2 == 'FOUR'):
                            ind_batter_stats[current_batter][2] += 1
                            pak_bowler_stats[current_bowler][2] += 4
                            ind_extras += 4
                        if(runs2 == '1 run'):
                            pak_bowler_stats[current_bowler][2] += 1
                            ind_extras += 1
                        if(runs2 == '2 runs'):
                            pak_bowler_stats[current_bowler][2] += 2
                            ind_extras += 2
                        if(runs2 == '3 runs'):
                            pak_bowler_stats[current_bowler][2] += 3
                            ind_extras += 3
                else:
                    runs = line[j:l]
                    if(runs == 'SIX'):
                        run = 6 
                        ind_batter_stats[current_batter][3] += 1
                        pak_bowler_stats[current_bowler][2] += 6
                    if(runs == 'FOUR'):
                        run = 4
                        ind_batter_stats[current_batter][2] += 1
                        pak_bowler_stats[current_bowler][2] += 4
                    if(runs == '1 run'):
                        run = 1
                        pak_bowler_stats[current_bowler][2] += 1
                    if(runs == '2 runs'):
                        run = 2
                        pak_bowler_stats[current_bowler][2] += 2
                    if(runs == '3 runs'):
                        run = 3
                        pak_bowler_stats[current_bowler][2] += 3
                    if(runs != 'wide'):
                        ind_batter_stats[current_batter][1] += 1
                    if(runs ==  'wide'):
                        pak_bowler_stats[current_bowler][2] += 1
                        pak_bowler_stats[current_bowler][5] += 1
                        ind_extras += 1
                    if(runs ==  '2 wides'):
                        pak_bowler_stats[current_bowler][2] += 2
                        pak_bowler_stats[current_bowler][5] += 2
                        ind_extras += 2
                    if(runs ==  '3 wides'):
                        pak_bowler_stats[current_bowler][2] += 3
                        pak_bowler_stats[current_bowler][5] += 3
                        ind_extras += 3
                    if((runs == 'leg byes') | (runs == 'byes')):
                        j = line.index(' ', k+1) + 1
                        k = line.index(',', k+1)
                        runs2 = line[j:k]
                        if(runs2 == 'SIX'):
                            ind_batter_stats[current_batter][3] += 1
                            pak_bowler_stats[current_bowler][2] += 6
                            ind_extras += 6
                        if(runs2 == 'FOUR'):          
                            ind_batter_stats[current_batter][2] += 1
                            pak_bowler_stats[current_bowler][2] += 4
                            ind_extras += 4
                        if(runs2 == '1 run'):
                            pak_bowler_stats[current_bowler][2] += 1
                            ind_extras += 1
                        if(runs2 == '2 runs'):
                            pak_bowler_stats[current_bowler][2] += 2
                            ind_extras += 2
                        if(runs2 == '3 runs'):
                            pak_bowler_stats[current_bowler][2] += 3
                            ind_extras += 3
                    if(runs[:3] == 'out'):
                        pak_bowler_stats[current_bowler][3] += 1
                        n_out += 1
                        now_runs = 0
                        for batter in ind_batter_stats:
                            now_runs += ind_batter_stats[batter][0]
                        now_runs += ind_extras
                        ind_batter_stats[current_batter][5] = f'{now_runs}-{n_out}'
                        p = line.index(' ')
                        ind_batter_stats[current_batter][6] = f'{line[:p]}'
                        if(runs == 'out Lbw'):
                            ind_batter_stats[current_batter][-1] = f'lbw {current_bowler}'
                        if(runs == 'out Bowled'):
                            ind_batter_stats[current_batter][-1] = f'b {current_bowler}'
                        if(runs[4:10] == 'Caught'):
                            caught_by = runs[14:l]
                            for player in team_pak:
                                if(caught_by in player):
                                    caught_by = player
                            ind_batter_stats[current_batter][-1] = f'c {caught_by} b {current_bowler}'
                ind_batter_stats[current_batter][0] += run
            except:
                pass
            
            if(current_ball == '5.6'):
                for batter in ind_batter_stats:
                    ind_power_play_runs += ind_batter_stats[batter][0]
                ind_power_play_runs += ind_extras

    for batter in ind_batter_stats:
        ind_batter_stats[batter][4] =  round(float(ind_batter_stats[batter][0]/ind_batter_stats[batter][1])*100,2)
    for bowler in pak_bowler_stats:
        last_over = pak_bowler_stats[bowler][0][-1]
        overs = len(pak_bowler_stats[bowler][0])
        if(type(last_over) == float):
            overs = round((overs + last_over - math.floor(last_over)),1)
        pak_bowler_stats[bowler][0].append(overs)
        num_overs = math.floor(pak_bowler_stats[bowler][0][-1]) + (pak_bowler_stats[bowler][0][-1] - math.floor(pak_bowler_stats[bowler][0][-1]))/0.6
        pak_bowler_stats[bowler][-1] =  round(float(pak_bowler_stats[bowler][2]/num_overs),1)

    ind_total = 0
    ind_outs = 0
    for batter in ind_batter_stats:
        ind_total += ind_batter_stats[batter][0]
        if ind_batter_stats[batter][-1]!='not out':
            ind_outs += 1
    ind_total += ind_extras

    ind_fall_of_wickets = []
    for batter in ind_batter_stats:
        if(ind_batter_stats[batter][-1] != 'not out'):
            fall = f'{ind_batter_stats[batter][-3]} ({batter}, {ind_batter_stats[batter][-2]})'
            ind_fall_of_wickets.append(fall)
    
    ind_fall_of_wickets.sort(key=get_fall)

    ind_score = str(ind_total) + '-' + str(ind_outs) + f' ({str(ind_last_over)} Ov)'
    print(f"\n{'## INDIA INNINGS': <18} {ind_score: <10}\n")
    print(f"{'Batter': <23}{' ': <45}{'R': ^5}{'B': ^5}{'4s': ^5}{'6s': ^5}{'SR': >7}")
    for batter in ind_batter_stats:
        print(f"{batter: <23}{ind_batter_stats[batter][-1]: <45}{ind_batter_stats[batter][0]: ^5}{ind_batter_stats[batter][1]: ^5}{ind_batter_stats[batter][2]: ^5}{ind_batter_stats[batter][3]: ^5}{ind_batter_stats[batter][4]: ^8}")
    print(f"{'Extras': <23}{'': <45}{ind_extras: ^5}")
    ind_total = f'{ind_total}({ind_outs} Wk, {ind_last_over} Ov)' 
    print(f"{'Total': <23}{'': <45}{ind_total: ^5}")

    print('\nFall of Wickets')
    fall_statement = ''
    for outs in ind_fall_of_wickets:
        fall_statement += outs + ', '
    print(fall_statement[:-2])

    print(f"\n{'Bowler': <23}{'O': ^5}{'M': ^5}{'R': ^5}{'W': ^5}{'NB': ^5}{'WD': ^5}{'ECO': >5}")
    for bowler in pak_bowler_stats:
        print(f"{bowler: <23}{pak_bowler_stats[bowler][0][-1]: ^5}{pak_bowler_stats[bowler][1]: ^5}{pak_bowler_stats[bowler][2]: ^5}{pak_bowler_stats[bowler][3]: ^5}{pak_bowler_stats[bowler][4]: ^5}{pak_bowler_stats[bowler][5]: ^5}{pak_bowler_stats[bowler][6]: ^7}")

    print(f"\n{'Powerplays': <15}{'Overs': ^8}{'Runs': >8}")
    print(f"{'Mandatory': <15}{'0.1-6': ^8}{ind_power_play_runs: >8}")

    #saving ind innings
    ind_innings_scorecard(ind_batter_stats, ind_score, ind_fall_of_wickets, pak_bowler_stats, ind_power_play_runs )

def pak_innings_scorecard(pak_batter_stats, pak_score, pak_fall_of_wickets, ind_bowler_stats, pak_power_play_runs):
    
    with open('Scorecard.txt','w') as scorecard:
        scorecard.write(f"{'## PAKISTAN INNINGS': <18} {pak_score: <10}\n")
        scorecard.write(f"\n{'Batter': <23}{' ': <45}{'R': ^5}{'B': ^5}{'4s': ^5}{'6s': ^5}{'SR': >7}")
        for batter in pak_batter_stats:
            scorecard.write(f"\n{batter: <23}{pak_batter_stats[batter][-1]: <45}{pak_batter_stats[batter][0]: ^5}{pak_batter_stats[batter][1]: ^5}{pak_batter_stats[batter][2]: ^5}{pak_batter_stats[batter][3]: ^5}{pak_batter_stats[batter][4]: ^8}")

        scorecard.write('\n\nFall of Wickets\n')
        fall_statement = ''
        for outs in pak_fall_of_wickets:
            fall_statement += outs + ', '
        scorecard.write(fall_statement[:-2])

        scorecard.write(f"\n\n{'Bowler': <23}{'O': ^5}{'M': ^5}{'R': ^5}{'W': ^5}{'NB': ^5}{'WD': ^5}{'ECO': >5}")
        for bowler in ind_bowler_stats:
            scorecard.write(f"\n{bowler: <23}{ind_bowler_stats[bowler][0][-1]: ^5}{ind_bowler_stats[bowler][1]: ^5}{ind_bowler_stats[bowler][2]: ^5}{ind_bowler_stats[bowler][3]: ^5}{ind_bowler_stats[bowler][4]: ^5}{ind_bowler_stats[bowler][5]: ^5}{ind_bowler_stats[bowler][6]: ^7}")

        scorecard.write(f"\n\n{'Powerplays': <15}{'Overs': ^8}{'Runs': >8}")
        scorecard.write(f"\n{'Mandatory': <15}{'0.1-6': ^8}{pak_power_play_runs: >8}")

def pak_innings(team_pak, team_ind):
    pak_extras=0
    ind_players = {'Batters': [], 'Bowlers': []}
    pak_players = {'Batters': [], 'Bowlers': []}
    with open('pak_inns1.txt') as inns1:
        for line in inns1:
            if (line == '\n'):
                continue
            j = line.index(' ') + 1
            k = line.index(',')
            delivery = line[j:k].split('to')

            for i in range(len(delivery)):
                delivery[i] = delivery[i].strip()
            for player in team_ind:
                if (delivery[0] in player):
                    if (player not in ind_players['Bowlers']):
                        ind_players['Bowlers'].append(player)
            for player in team_pak:
                if (delivery[1] in player):
                    if (player not in pak_players['Batters']):
                        pak_players['Batters'].append(player)

        pak_batter_stats = {}
        ind_bowler_stats = {}

        for batter in pak_players['Batters']:
            pak_batter_stats[batter] = [0]*8
            pak_batter_stats[batter][-1] = 'not out'

        for bowler in ind_players['Bowlers']:
            ind_bowler_stats[bowler] = [[]]
            for i in range(6):
                ind_bowler_stats[bowler].append(0)
        last_line = line
        pak_last_over = last_line[:last_line.index('.')+2]

    with open('pak_inns1.txt') as inns1:
            i = last_line.index('.')
            last_over = float(line[:i+2]) + 1
            n_out = 0
            pak_power_play_runs = 0
            for line in inns1:
                if (line == '\n'):
                    continue
                i = line.index('.')
                current_over = int(line[:i]) + 1
                current_ball = line[:i+2]
                try:
                    j = line.index(' ') + 1
                    k = line.index(',')
                    delivery = line[j:k].split('to')
                    for i in range(len(delivery)):
                        delivery[i] = delivery[i].strip()
                    for player in team_ind:
                        if (delivery[0] in player):
                            current_bowler = player
                    for player in team_pak:
                        if (delivery[1] in player):
                            current_batter = player
                except:
                    k = 0
                if (current_over not in ind_bowler_stats[f'{current_bowler}'][0]):
                    if (current_over == math.floor(last_over)):
                        last_over = float(line[:i+3]) + 1
                        ind_bowler_stats[current_bowler][0][-1] = last_over
                    else:
                        ind_bowler_stats[current_bowler][0].append(current_over)
                try:
                    try:
                        j = line.index(' ', k+1) + 1
                        k = line.index(',', k+1)
                    except:
                        k = 10000000
                    try:
                        l = line.index('!')
                    except:
                        l = 10000000
                    run = 0
                    if (k < l):
                        runs = line[j:k]
                        if (runs == 'SIX'):
                            run = 6
                            pak_batter_stats[current_batter][3] += 1
                            ind_bowler_stats[current_bowler][2] += 6
                        if (runs == 'FOUR'):
                            run = 4
                            pak_batter_stats[current_batter][2] += 1
                            ind_bowler_stats[current_bowler][2] += 4
                        if (runs == '1 run'):
                            run = 1
                            ind_bowler_stats[current_bowler][2] += 1
                        if (runs == '2 runs'):
                            run = 2
                            ind_bowler_stats[current_bowler][2] += 2
                        if (runs == '3 runs'):
                            run = 3
                            ind_bowler_stats[current_bowler][2] += 3
                        if (runs != 'wide'):
                            pak_batter_stats[current_batter][1] += 1
                        if (runs == 'wide'):
                            ind_bowler_stats[current_bowler][2] += 1
                            ind_bowler_stats[current_bowler][5] += 1
                            pak_extras += 1
                        if (runs == '2 wides'):
                            ind_bowler_stats[current_bowler][5] += 2
                            pak_extras += 2
                        if (runs == '3 wides'):
                            ind_bowler_stats[current_bowler][5] += 3
                            pak_extras += 3
                        if ((runs == 'leg byes') | (runs == 'byes')):
                            j = line.index(' ', k+1) + 1
                            k = line.index(',', k+1)
                            runs2 = line[j:k]
                            if (runs2 == 'SIX'):
                                pak_batter_stats[current_batter][3] += 1
                                ind_bowler_stats[current_bowler][2] += 6
                                pak_extras += 6
                            if (runs2 == 'FOUR'):
                                pak_batter_stats[current_batter][2] += 1
                                ind_bowler_stats[current_bowler][2] += 4
                                pak_extras += 4
                            if (runs2 == '1 run'):
                                ind_bowler_stats[current_bowler][2] += 1
                                pak_extras += 1
                            if (runs2 == '2 runs'):
                                ind_bowler_stats[current_bowler][2] += 2
                                pak_extras += 2
                            if (runs2 == '3 runs'):
                                ind_bowler_stats[current_bowler][2] += 3
                                pak_extras += 3
                    else:
                        runs = line[j:l]
                        if (runs == 'SIX'):
                            run = 6
                            pak_batter_stats[current_batter][3] += 1
                            ind_bowler_stats[current_bowler][2] += 6
                        if (runs == 'FOUR'):
                            run = 4
                            pak_batter_stats[current_batter][2] += 1
                            ind_bowler_stats[current_bowler][2] += 4
                        if (runs == '1 run'):
                            run = 1
                            ind_bowler_stats[current_bowler][2] += 1
                        if (runs == '2 runs'):
                            run = 2
                            ind_bowler_stats[current_bowler][2] += 2
                        if (runs == '3 runs'):
                            run = 3
                            ind_bowler_stats[current_bowler][2] += 3
                        if (runs != 'wide'):
                            pak_batter_stats[current_batter][1] += 1
                        if (runs == 'wide'):
                            ind_bowler_stats[current_bowler][2] += 1
                            ind_bowler_stats[current_bowler][5] += 1
                            pak_extras += 1
                        if (runs == '2 wides'):
                            ind_bowler_stats[current_bowler][2] += 2
                            ind_bowler_stats[current_bowler][5] += 2
                            pak_extras += 2
                        if (runs == '3 wides'):
                            ind_bowler_stats[current_bowler][2] += 3
                            ind_bowler_stats[current_bowler][5] += 3
                            pak_extras += 3
                        if ((runs == 'leg byes') | (runs == 'byes')):
                            j = line.index(' ', k+1) + 1
                            k = line.index(',', k+1)
                            runs2 = line[j:k]
                            if (runs2 == 'SIX'):
                                pak_batter_stats[current_batter][3] += 1
                                ind_bowler_stats[current_bowler][2] += 6
                                pak_extras += 6
                            if (runs2 == 'FOUR'):
                                pak_batter_stats[current_batter][2] += 1
                                ind_bowler_stats[current_bowler][2] += 4
                                pak_extras += 4
                            if (runs2 == '1 run'):
                                ind_bowler_stats[current_bowler][2] += 1
                                pak_extras += 1
                            if (runs2 == '2 runs'):
                                ind_bowler_stats[current_bowler][2] += 2
                                pak_extras += 2
                            if (runs2 == '3 runs'):
                                ind_bowler_stats[current_bowler][2] += 3
                                pak_extras += 3
                        if (runs[:3] == 'out'):
                            ind_bowler_stats[current_bowler][3] += 1
                            n_out += 1
                            now_runs = 0
                            for batter in pak_batter_stats:
                                now_runs += pak_batter_stats[batter][0]
                            now_runs += pak_extras
                            pak_batter_stats[current_batter][5] = f'{now_runs}-{n_out}'
                            p = line.index(' ')
                            pak_batter_stats[current_batter][6] = f'{line[:p]}'
                            if (runs == 'out Lbw'):
                                pak_batter_stats[current_batter][-1] = f'lbw {current_bowler}'
                            if (runs == 'out Bowled'):
                                pak_batter_stats[current_batter][-1] = f'b {current_bowler}'
                            if (runs[4:10] == 'Caught'):
                                caught_by = runs[14:l]
                                for player in team_ind:
                                    if (caught_by in player):
                                        caught_by = player
                                pak_batter_stats[current_batter][-1] = f'c {caught_by} b {current_bowler}'
                    pak_batter_stats[current_batter][0] += run
                except:
                    pass
                if (current_ball == '5.6'):
                    for batter in pak_batter_stats:
                        pak_power_play_runs += pak_batter_stats[batter][0]
                    pak_power_play_runs += pak_extras

    for batter in pak_batter_stats:
        pak_batter_stats[batter][4] =  round(float(pak_batter_stats[batter][0]/pak_batter_stats[batter][1])*100,2)

    for bowler in ind_bowler_stats:
            last_over = ind_bowler_stats[bowler][0][-1]
            overs = len(ind_bowler_stats[bowler][0])
            if(type(last_over) == float):
                overs = round((overs + last_over - math.floor(last_over)),1)
            ind_bowler_stats[bowler][0].append(overs)
            num_overs = math.floor(ind_bowler_stats[bowler][0][-1]) + (ind_bowler_stats[bowler][0][-1] - math.floor(ind_bowler_stats[bowler][0][-1]))/0.6
            ind_bowler_stats[bowler][-1] =  round(float(ind_bowler_stats[bowler][2]/num_overs),1)

    pak_total, pak_outs = 0, 0
    for batter in pak_batter_stats:
        pak_total += pak_batter_stats[batter][0]
        if pak_batter_stats[batter][-1]!='not out':
            pak_outs+=1
    pak_total += pak_extras

    pak_fall_of_wickets = []
    for batter in pak_batter_stats:
        if(pak_batter_stats[batter][-1] != 'not out'):
            fall = f'{pak_batter_stats[batter][-3]} ({batter}, {pak_batter_stats[batter][-2]})'
            pak_fall_of_wickets.append(fall)
    
    pak_fall_of_wickets.sort(key=get_fall)

    #printing scorecard
    pak_score = str(pak_total) + '-' + str(pak_outs) + f' ({str(pak_last_over)} Ov)'
    print(f"{'## PAKISTAN INNINGS': <18} {pak_score: <10}\n")
    print(f"{'Batter': <23}{' ': <45}{'R': ^5}{'B': ^5}{'4s': ^5}{'6s': ^5}{'SR': >7}")
    for batter in pak_batter_stats:
        print(f"{batter: <23}{pak_batter_stats[batter][-1]: <45}{pak_batter_stats[batter][0]: ^5}{pak_batter_stats[batter][1]: ^5}{pak_batter_stats[batter][2]: ^5}{pak_batter_stats[batter][3]: ^5}{pak_batter_stats[batter][4]: ^8}")

    pak_total = f'{pak_total}({pak_outs} Wk, {pak_last_over} Ov)' 
    print(f"{'Extras': <23}{'': <45}{pak_extras: ^5}")
    print(f"{'Total': <23}{'': <45}{pak_total: ^5}")

    print('\nFall of Wickets')
    fall_statement = ''
    for outs in pak_fall_of_wickets:
        fall_statement += outs + ', '
    print(fall_statement[:-2])

    print(f"\n{'Bowler': <23}{'O': ^5}{'M': ^5}{'R': ^5}{'W': ^5}{'NB': ^5}{'WD': ^5}{'ECO': >5}")
    for bowler in ind_bowler_stats:
        print(f"{bowler: <23}{ind_bowler_stats[bowler][0][-1]: ^5}{ind_bowler_stats[bowler][1]: ^5}{ind_bowler_stats[bowler][2]: ^5}{ind_bowler_stats[bowler][3]: ^5}{ind_bowler_stats[bowler][4]: ^5}{ind_bowler_stats[bowler][5]: ^5}{ind_bowler_stats[bowler][6]: ^7}")
    
    print(f"\n{'Powerplays': <15}{'Overs': ^8}{'Runs': >8}")
    print(f"{'Mandatory': <15}{'0.1-6': ^8}{pak_power_play_runs: >8}")

    #saving pak innings
    pak_innings_scorecard(pak_batter_stats, pak_score, pak_fall_of_wickets, ind_bowler_stats, pak_power_play_runs )

def team_ind_list():
    with open('teams.txt') as team_file:
        team_ind_str=''
        for line in team_file:
            if line[0]=='I':
                team_ind_str=line

        start_index_ind=team_ind_str.find(':')
        ind_team_xi=team_ind_str[start_index_ind+2:len(team_ind_str)-1]
        team_ind=ind_team_xi.split(', ')

        return team_ind

def team_pak_list():
    with open('teams.txt') as team_file:
        team_pak_str=''
        for line in team_file:
            if line[0]=='P':
                team_pak_str=line

        start_index_pak=team_pak_str.find(':')
        pak_team_xi=team_pak_str[start_index_pak+2:len(team_pak_str)-1]
        team_pak=pak_team_xi.split(', ')

        return team_pak

def scorecard():
    pak_extras = 0
    ind_extras = 0
    team_pak = team_pak_list()
    team_ind = team_ind_list()
    pak_innings(team_pak, team_ind)
    ind_innings(team_pak, team_ind)

scorecard()

#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))