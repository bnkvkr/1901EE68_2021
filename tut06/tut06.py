import re
import os
from os import listdir
from distutils.dir_util import copy_tree


def pre_process():
    renamed_srt = './corrected_srt'

    isfile2 = os.path.isdir(renamed_srt)
    if(isfile2 == False):
        os.mkdir(renamed_srt)


def BB(season_no, episode_no, series_name):
    mypath = f'./corrected_srt/{series_name}'
    wrong_path = f'./wrong_srt/{series_name}'
    is_break_path = os.path.isdir(mypath)
    if(is_break_path == False):
        os.mkdir(mypath)
        copy_tree(wrong_path, mypath)
    else:
        for f in os.listdir(mypath):
            os.remove(os.path.join(mypath, f))
        copy_tree(wrong_path, mypath)
    from os.path import isfile, join

    onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]

    for onlyfile in onlyfiles:
        text = onlyfile
        pattern = re.compile(r'\d+')
        found = re.findall(pattern, text)
        season = found[0].lstrip('0')
        episode = found[1].lstrip('0')
        for i in range(season_no - len(season)):
            season = '0' + season

        for i in range(episode_no - len(episode)):
            episode = '0' + episode

        if onlyfile[-3:] == 'srt':
            old_file_name = f"./corrected_srt/{series_name}/{onlyfile}"
            new_file_name = f"./corrected_srt/{series_name}/{series_name} Season {season} Episode {episode}.srt"
            os.rename(old_file_name, new_file_name)
        else:
            old_file_name = f"./corrected_srt/{series_name}/{onlyfile}"
            new_file_name = f"./corrected_srt/{series_name}/{series_name} Season {season} Episode {episode}.mp4"

            os.rename(old_file_name, new_file_name)


def GOT(season_no, episode_no, series_name):
    mypath = f'./corrected_srt/{series_name}'
    wrong_path = f'./wrong_srt/{series_name}'
    is_break_path = os.path.isdir(mypath)
    if(is_break_path == False):
        os.mkdir(mypath)
        copy_tree(wrong_path, mypath)
    else:
        for f in os.listdir(mypath):
            os.remove(os.path.join(mypath, f))
        copy_tree(wrong_path, mypath)

    from os.path import isfile, join
    onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]
    for onlyfile in onlyfiles:
        text = onlyfile

        pattern = re.compile(r'\d+')
        episode_name = onlyfile.split('-')
        episode_name = episode_name[2].split('.')[0]
        found = re.findall(pattern, text)
        season = found[0].lstrip('0')
        episode = found[1].lstrip('0')
        for i in range(season_no - len(season)):
            season = '0' + season
        for i in range(episode_no - len(episode)):
            episode = '0' + episode
        if onlyfile[-3:] == 'srt':
            old_file_name = f"./corrected_srt/{series_name}/{onlyfile}"
            new_file_name = f"./corrected_srt/{series_name}/{series_name} - Season {season} Episode {episode} -{episode_name}.srt"
            os.rename(old_file_name, new_file_name)
        else:
            old_file_name = f"./corrected_srt/{series_name}/{onlyfile}"
            new_file_name = f"./corrected_srt/{series_name}/{series_name} - Season {season} Episode {episode} -{episode_name}.mp4"
            os.rename(old_file_name, new_file_name)


def regex_renamer():
    webseries_num = int(
        input("Enter the number of the web series that you wish to rename. 1/2/3: "))
    season_padding = int(input("Enter the Season Number Padding: "))
    episode_padding = int(input("Enter the Episode Number Padding: "))
    if webseries_num == 1:
        series_name = 'Breaking Bad'
        BB(season_padding, episode_padding, series_name)
    if webseries_num == 2:
        series_name = 'Game of Thrones'
        GOT(season_padding, episode_padding, series_name)
    if webseries_num == 3:
        series_name = 'Lucifer'
        GOT(season_padding, episode_padding, series_name)
    return


pre_process()
regex_renamer()
