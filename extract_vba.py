#-*- coding: utf-8 -*-
import math
import sys
import os
import zlib
import olevba
import time
reload(sys)
sys.setdefaultencoding('utf8')

totalcount = {}
def Extract(root, name):
	#olevba.process_file(None, sys.argv[1], None, show_decoded_strings=None)
	try:
		vba = olevba.VBA_Parser(os.path.join(root,name), None)
		if vba.detect_vba_macros():
			try:
				extracted_macros = vba.extract_macros()
			except:
				return
			i = 0
			for (subfilename, stream_path, vba_filename, vba_code) in extracted_macros:
				vba_code_filtered = olevba.filter_vba(vba_code)
				# detect empty macros
				if vba_code_filtered.strip() == '':
					#print '(empty macro)'
					continue
				else:
					try:
						fileaddr = os.path.join(root,name+"+"+vba_filename+".vb")
						resfile = open(fileaddr, "w")
					except IOError:
						fileaddr = os.path.join(root,name+"+"+str(i)+".vb")
						resfile = open(fileaddr, "w")
						i = i+1
					resfile.write(vba_code_filtered)
					resfile.close()
		# else:
			# print 'No VBA macros found.'
	except:
		return
		
def main():
	if len(sys.argv)!=2:
		print "usage: ", os.path.basename(sys.argv[0]), " <filename or dirname>\n"
		sys.exit()
	for root, dirs, files in os.walk(sys.argv[1]):
		for name in files:
			if name.endswith(".vb"):
				continue
			elif name.endswith(".zip"):
				continue
			elif name.endswith(".txt"):
				continue
			else:
				Extract(root, name)
if __name__ == '__main__':
    main()