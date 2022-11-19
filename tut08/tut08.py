import os
import math
from datetime import datetime
start_time = datetime.now()



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
    # pak_innings(team_pak, team_ind, pak_extras)

scorecard()

#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))