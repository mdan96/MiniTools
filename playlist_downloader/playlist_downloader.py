from pytube import YouTube
from pytube import Playlist
import os 
import re

def create_playlist_dict():
    playlist_dict = {}
    with open('links.txt', 'r') as file:
        lines = file.readlines()
        for line in lines:
            if line.startswith('['):
                pass
            else:
                content = line.split(',')
                link = content[0]
                folder_name = content[1].replace('\n','')
                playlist_dict[link] = folder_name
    return playlist_dict

playlist_dict = create_playlist_dict()
for key in playlist_dict.keys():
    link = key
    folder_name = playlist_dict[key]
    playlist = Playlist(key)
    playlist.video_urls
    parent_dir = os.getcwd()
    path = os.path.join(parent_dir, folder_name)
    count = 0
    for url in playlist:
        count += 1
        print(count,' ',url)
        YouTube(url).streams.filter(only_audio=True).first().download(path)