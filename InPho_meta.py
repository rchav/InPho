import shutil
import os

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from mutagen.flac import FLAC

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





def buildMetaDataFrame(pathOfFiles):
	fullPaths, justFileNames = find_flacs(pathOfFiles) 

	df = pd.DataFrame(columns = ['Inmate Name', 'NYSID', 'BAC', 'FileName', 'NumberDialed', 
								'extension', 'facility', 'start_time (hhmmss)', 
								'start_date (yyyymmdd)'])
	
	# somewhere to store lists after the loop
	called_numbers = []
	nysids = []
	names = [] 
	extensions = []
	facilities = [] 
	bacs = []
	start_times = []
	start_dates = []

	# loop to get my list of values
	for file in fullPaths:
		track = FLAC(file)
		try:
			called_numbers .append(track['called_number'])
			nysids.append(track['nysid'])
			names.append(track['name'])
			extensions.append(track['extension'])
			facilities.append(track['facility'])
			bacs.append(track['bac'])
			start_times.append(track['start_time (hhmmss)'])
			start_dates.append(track['start_date (yyyymmdd)'])

		except Exception as err:
			print (err)

	# put the values into my pandas dataframe
	df['NumberDialed'] = called_numbers
	df['NYSID'] = nysids
	df['Inmate Name'] = names
	df['extension'] = extensions
	df['facility'] = facilities
	df['BAC'] = bacs
	df['start_time (hhmmss)'] = start_times
	df['start_date (yyyymmdd)'] = start_dates
	df['FileName'] = justFileNames

	numberOfFiles = len(fullPaths)

	result = df
	return result, numberOfFiles, fullPaths, justFileNames



def InPhoFrame(src):
	result, numberOfFiles, fullPaths, justFileNames = buildMetaDataFrame(src)

	print result



InPhoFrame(src)
