#!/usr/bin/python2
import re
from pprint import pprint
import errno
def open_file(filename):
	try:
		fp = open(filename)
	except IOError as e:
		if e.errno == errno.ENOENT:
			print("Error: Cannot find file {} at src/".format(filename))
			return None
		raise
	else:
		with fp:
			return fp.readlines()
def read_info(filename):
	lines = open_file(filename)
	if lines is None:
		return None
	infos = dict()
	for line in lines:
		if line.startswith('#') or line.isspace(): #skip comments
			continue
		m = re.search(r'(.+)="(.+)"', line)
		if m:
			infos[m.group(1)] = m.group(2)
		else:
			print("Error: formatting errors detected. Is your {} correctly formatted?".format(filename))
			return None
	return infos
def main():
	infos = read_info("user_info.config")
	pprint(infos)
if __name__ == "__main__":
	main()
