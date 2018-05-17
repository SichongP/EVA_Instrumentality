#!/usr/bin/python2
import os
import os.path
import argparse
import project
import writexl
import user
from pprint import pprint
from datetime import date
#usage: mkmetadata [options] path_to_project_folder
#Options:
#--use-user-info user_info.config	use config file to get user information

parser = argparse.ArgumentParser(description = "Generate metadata file based on input files")
#parser.add_argument('--user', type=argparse.FileType('r'), default=None)
parser.add_argument('--user', default=None)
parser.add_argument('path', default="./projects")
parser.add_argument('--out', default= "autogenerated_EVA_submission" + str(date.today()) + ".xlsx")
#a = parser.parse_args(['--user', 'user_info.config'])
a = parser.parse_args()
read_user = False
user_infos = None
if os.path.isfile(a.out):
	print("Error! output file {} already exists! Use another output file name!".format(a.out))
if a.user:
	print(a.user)
	user_infos = user.read_info(a.user)
	read_user = True
	pprint(user_infos)
projects = []
for dir in os.listdir(a.path):
	print(dir)
	projects.append(project.project(os.path.join(a.path,dir)))
writexl.write(read_user, user_infos, projects, a.out)