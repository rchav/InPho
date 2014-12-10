import shutil
import os
import pandas
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


def flac_metadata_reader(path):
	fullPaths, justFileNames = find_flacs(path)

	for file in fullPaths:
		track = FLAC(file)
		try:
			print track
			print '\n'
		except Exception as err:
			print (err)

	print str(len(justFileNames)) + " files in this directory \n"


flac_metadata_reader(src)
