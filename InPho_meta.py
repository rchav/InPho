import shutil
import os

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from datetime import datetime

from mutagen.flac import *


src = '/Projects/InPho/testing/'

def find_flacs(path):

    # create some lists of flac files and their paths
    fullPaths = []
    justFileNames = []

    for root, dirs, files in os.walk(path):
        for file in files:

            # create a list of the full paths
            fullFilePath = os.path.join(root, file)
            if fullFilePath.endswith('.flac'):
              fullPaths.append(fullFilePath)

            # create a list of the flac file names
            justFlacFileName = os.path.join(file)
            if justFlacFileName.endswith('.flac'):
                justFileNames.append(justFlacFileName)

    return fullPaths, justFileNames


def build_metadata_frame(pathOfFiles):
  fullPaths, justFileNames = find_flacs(pathOfFiles) 

  df = pd.DataFrame(columns = ['InmateName', 'NYSID', 'BAC', 'DateTime', 'NumberDialed', 
             'Duration', 'Facility', 'Extension', 'FileName', 'start_time', 
             'start_date'])
  
  # somewhere to store lists after the loop
  called_numbers = []
  nysids = []
  names = [] 
  extensions = []
  facilities = [] 
  bacs = []
  start_times = []
  start_dates = []
  datetimes = []
  duration = []

  # loop to get my list of values
  for file in fullPaths:
    track = FLAC(file)
    try:
      called_numbers.append(''.join(track.get('called_number')))
      nysids.append(''.join(track.get('nysid')))
      names.append(''.join(track.get('name')))
      extensions.append(''.join(track.get('extension')))
      facilities.append(''.join(track.get('facility')))
      bacs.append(''.join(track.get('bac')))
      start_times.append(''.join(track.get('start_time (hhmmss)')))
      start_dates.append(''.join(track.get('start_date (yyyymmdd)')))
      duration.append(track.info.length)
      datetimes.append(str(''.join(track.get('start_date (yyyymmdd)'))) + str(''.join(track.get('start_time (hhmmss)'))))

    except Exception as err:
      print (err)


  # put the values into my pandas dataframe
  df['NumberDialed'] = called_numbers
  df['NYSID'] = nysids
  df['InmateName'] = names
  df['Extension'] = extensions
  df['Facility'] = facilities
  df['BAC'] = bacs
  df['start_time'] = start_times
  df['start_date'] = start_dates
  df['FileName'] = justFileNames
  df['Duration'] = duration


  numberOfFiles = len(fullPaths)

  result = df
  return result, numberOfFiles, fullPaths, justFileNames



def create_InPho_frame(src):

  # take in a pandas dataframe after scanning for flacs
  result, numberOfFiles, fullPaths, justFileNames = build_metadata_frame(src)
  df = result 

  # toss the isolated time and date columns
  df['DateTime'] = pd.DatetimeIndex(df['start_date'] + df['start_time'])
  df = df.drop(['start_time', 'start_date'], 1)

  result = df

  return result
