#!/usr/bin/python2

#This module defines project, analysis, and file calsses
import os
import os.path
from user import open_file, read_info
from pprint import pprint
import hashlib
import re

class file:
	def get_md5(self, filename):
		hash_md5 = hashlib.md5()
		with open(filename, "rb") as f:
			for chunk in iter(lambda: f.read(131072), b""):
				hash_md5.update(chunk)
				return hash_md5.hexdigest()
	def print_file_info(self):
		print(self.name)
		print(self.filetype)
		print(self.md5)
	def get_filetype(self, filename):
		#check key word in file name to determine file type
		#Only specified file types are allowed so only check these key words
		if re.search(r'\.vcf', filename):
			return "vcf"
		elif re.search(r'\.vcf.+aggregate|aggregate.+\.vcf', filename):
			return "vcf_aggregate"
		elif re.search(r'\.cram', filename):
			return "cram"
		elif re.search(r'\.tabix', filename):
                        return "tabix"
		elif re.search(r'\.wig', filename):
                        return "wig"
		elif re.search(r'\.bed', filename):
                        return "bed"
		elif re.search(r'\.gff', filename):
                        return "gff"
		elif re.search(r'\.fasta|\.fa|\.fna', filename):
                        return "fasta"
		elif re.search(r'read_me|readme|README|READ_ME', filename):
			return "readme_file"
		else:
			return "other"
	def __init__(self, anly, path):
	#a path to the file is expected
		self.analysis = anly
		(temp, self.name) = os.path.split(path)
		self.filetype = self.get_filetype(self.name)
		self.md5 = self.get_md5(path)
	
class analysis:
	def read_analysis_info(self, path):
		anl_infos = read_info(path)
		if anl_infos:
			self.title = anl_infos['TITLE']
			self.description = anl_infos['DESCRIPTION']
			self.alias = anl_infos['ALIAS']
			self.experiment = anl_infos['EXPERIMENT_TYPE']
			self.ref = anl_infos['REFERENCE']
			self.refmd5 = anl_infos['REFMD5']
	def print_analy_info(self):
		print("Analysis title: {}".format(self.title))
		print("Analysis alias: {}".format(self.alias))
		print("Analysis description: {}".format(self.description))
		print("Experiment Type: {}".format(self.experiment))
		print("Reference: {}".format(self.ref))
		print("Associated files:\n\t\t")
		for item in self.files:
			item.print_file_info()
	def __init__(self, proj, analysis_count, path):
	#a path to the analysis directory is expected
		self.title = None
		self.alias = None
		self.description = None
		self.experiment = None
		self.ref = None
		self.refmd5 = None
		self.prj = proj
		self.files = []
		self.anly_info_loaded = False
		for dir in os.listdir(path):
			if dir == "analysis_info.config":
			#project_info.config file detected, read info from file
				self.read_analysis_info(os.path.join(path,dir))
				self.anly_info_loaded = True
			elif os.path.isfile(os.path.join(path,dir)):
				self.files.append(file(self, os.path.join(path,dir)))
			else:
				print("Unrecognized file {}, ignoring...".format(dir))
		if not self.anly_info_loaded:
			print("No analysis info file found in {}, be sure you input them manually".format(path))
		if not self.title:
		#analysis title not provided by user, read from project folder name instead
			if path.endswith(os.sep):
				path = path.rstrip(os.sep) #remove trailing slash if existed so we can get project fo$                        (temp, self.title) = os.path.split()
			(temp, self.title) = os.path.split(path)
		if not self.alias:
		#analysis alias not provided by user, auto-generating...
			if proj.alias:
				self.alias = proj.alias + "_" + str(analysis_count)
			else:
				self.alias = proj.title + "_" + str(analysis_count)
class project:
	def read_proj_info(self, path):
		prj_infos = read_info(path)
		if prj_infos:
			self.title = prj_infos['TITLE']
			self.alias = prj_infos['ALIAS']
			self.description = prj_infos['DESCRIPTION']
			self.center = prj_infos['CENTER']
			self.taxid = prj_infos['TAXID']
	def print_proj_info(self):
		print("Project title: {}".format(self.title))
		print("Project alias: {}".format(self.alias))
		print("Project description: {}".format(self.description))
		print("Project center: {}".format(self.center))
		print("Project taxid: {}".format(self.taxid))
		print("Project analyses:\n\t\t")
		for item in self.analyses:
			item.print_analy_info()
	def __init__(self, path):
	#a path to the project directory is expected
		self.title = None
		self.alias = None
		self.description = None
		self.center = None
		self.taxid = None
		#for each analysis folder, initiate an anlysis instance
		self.analyses = []
		self.proj_info_loaded = False
		self.analysis_count = 0 #record the how many analyses are encountered in case users don't provide analysis alias
		for dir in os.listdir(path):
			if dir == "project_info.config":
			#Look for project_info.config file, if detected, read info from file
				self.read_proj_info(os.path.join(path,dir))
				self.proj_info_loaded = True
		if not self.proj_info_loaded:
			print("No project info file found in {}, be sure you input them manually".format(path))
		if not self.title:
		#project title not provided by user, read from project folder name instead
			if path.endswith(os.sep):
				path = path.rstrip(os.sep) #remove trailing slash if existed so we can get project folder name
			(temp, self.title) = os.path.split(path)
		for dir in os.listdir(path):
			if dir == "project_info.config":
				continue
			if os.path.isdir(os.path.join(path,dir)):
				self.analysis_count += 1
				self.analyses.append(analysis(self, self.analysis_count, os.path.join(path,dir)))
			else:
				print("Unrecognized file {}, ignoring...".format(dir))
def main():
	a = project("/mnt/c/Users/pengs/GoogleDrive/FinnoLab/git_repo/EVA_Instrumentality/sample_structure/equine_NAD_CGAS/")
	a.print_proj_info()

if __name__ == "__main__":
        main()
