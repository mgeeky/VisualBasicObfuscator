#!/usr/bin/python

import re
import os
import sys
import string
import random
import argparse

config = {
	'verbose' : False,
	'quiet' : False,
	'file' : '',
	'output' : '',
	'garbage_perc' : 12.0,
	'min_var_length' : 5,
	'custom_reserved_words': []
}

def out(x):
	if not config['quiet']:
		sys.stderr.write(x + '\n')

def log(x):
	if config['verbose']:
		out(x)

def err(x):
	out('[!] ' + x)

def info(x):
	log('[?] ' + x)

def ok(x):
	log('[+] ' + x)

class ScriptObfuscator():

	RESERVED_NAMES = (
		"AutoExec",
		"AutoOpen",
		"DocumentOpen",
		"AutoExit",
		"AutoClose",
		"Document_Close",
		"DocumentBeforeClose",
		"Document_Open",
		"Document_BeforeClose",
		"Auto_Open",
		"Workbook_Open",
		"Workbook_Activate",
		"Auto_Close",
		"Workbook_Close"
	)

	def __init__(self, reserved_words, garbage_perc, min_var_length):
		self.input = ''
		self.output = ''
		self.garbage_perc = garbage_perc
		self.min_var_length = min_var_length
		self.reserved_words = reserved_words

	def obfuscate(self, inp):
		self.input = inp
		self.output = inp

		# Step 1: Remove comments
		self.output = re.sub(r"(?<!\"[^\"])'(.*)", "", self.output, flags=re.I)

		# Step 2: Rename used variables
		self.randomizeVariablesAndFunctions()

		# Step 3: Explode string constants
		self.obfuscateStrings()

		# Step 4: Obfuscate arrays
		self.obfuscateArrays()

		# Step 5: Insert garbage
		self.insertGarbage()

		# Step 6: Remove empty lines
		self.output = '\n'.join(filter(lambda x: not re.match(r'^\s*$', x), self.output.split('\n')))
		
		# Step 7: Remove indents and multi-spaces.
		self.output = re.sub(r"\t| {2,}", "", self.output, flags=re.I)

		return self.output

	def randomizeVariablesAndFunctions(self):
		# Variables
		for m in re.finditer(r"(?:(?:Dim|Set|Const)\s+(\w+)\s*(?:As|=)?)|(?:^(\w+)\s+As\s+)", self.output, flags = re.I|re.M):
			varName = randomString(random.randint(4,12))
			varToReplace = filter(lambda x: x and len(x)>0, m.groups())[0]

			if len(varToReplace) < self.min_var_length: continue
			if varToReplace in self.reserved_words: continue
			info("Variable name obfuscated (context: \"%s\"): '%s' => '%s'" % (m.group(0).strip(), varToReplace, varName))
			self.output = re.sub(r"\b" + varToReplace + r"\b", varName, self.output, flags=re.I | re.M)

		# Globals
		for m in re.finditer(r"\s*(?:(?:Public|Private|Protected)\s*(?:Dim|Set|Const)?\s+(\w+)\s*As)|(?:(?:Private|Protected|Public)\s+Declare\s+PtrSafe?\s*Function\s+(\w+)\s+)", self.output, flags = re.I|re.M):
			varName = randomString(random.randint(4,12))
			varToReplace = filter(lambda x: x and len(x)>0, m.groups())[0]

			if varToReplace in self.reserved_words: continue
			if len(varToReplace) < self.min_var_length: continue
			info("Global name obfuscated (context: \"%s\"): '%s' => '%s'" % (m.group(0).strip(), varToReplace, varName))
			self.output = re.sub(r"\b" + varToReplace + r"\b", varName, self.output, flags=re.I | re.M)			

		# Functions
		for m in re.finditer(r"\s*(?:Public|Private|Protected|Friend)?\s*(?:Sub|Function)\s+(\w+)\s*\(", self.output, flags = re.I|re.M):
			varName = randomString(random.randint(4,12))
			varToReplace = m.group(1)
			if len(varToReplace) < self.min_var_length: continue
			if varToReplace in self.reserved_words: continue
			if varToReplace in ScriptObfuscator.RESERVED_NAMES: continue
			info("Function name obfuscated (context: \"%s\"): '%s' => '%s'" % (m.group(0).strip(), varToReplace, varName))
			self.output = re.sub(r"\b" + varToReplace + r"\b", varName, self.output, flags=re.I | re.M)

	def obfuscateNumber(self, num):
		rnd1 = random.randint(0, 3333)
		num_coders = (
			lambda rnd1, num: '%d' % num,
			lambda rnd1, num: '%d-%d' % (rnd1+num, rnd1),
			lambda rnd1, num: '%d-%d' % (2*rnd1+num, 2*rnd1),
			lambda rnd1, num: '%d-%d' % (3*rnd1+num, 3*rnd1),
			lambda rnd1, num: '%d+%d' % (-rnd1, rnd1+num),
			lambda rnd1, num: '%d/%d' % (rnd1*num, rnd1),
		)

		return random.choice(num_coders)(rnd1, num)

	def obfuscateChar(self, char):
		char_coders = (
			lambda x: '"{}"'.format(x),
			lambda x: 'Chr(&H%x)' % ord(x),
			lambda x: 'Chr(%d)' % ord(x),
			lambda x: '"{}"'.format(x),
			lambda x: 'Chr(%s)' % self.obfuscateNumber(ord(x)),
			lambda x: 'Chr(Int("&H%x"))' % ord(x),
			lambda x: 'Chr(Int("%d"))' % ord(x),
		)
		out = random.choice(char_coders)(char)
		if out == '"""': out = '"\""'
		return out

	def obfuscateString(self, string):
		if len(string) == 0: return ""
		new_string = ''
		for char in string:
			new_string += self.obfuscateChar(char) + '&'
		return new_string[:-1]

	def obfuscateStrings(self):
		for m in re.finditer(r"(\"[^\"]+\"|\"\")", self.output, flags=re.I|re.M):
			orig_string = m.group(1)
			info("String to obfuscate (context: \"%s\"): %s" % (m.group(0).strip().replace('\n',''), orig_string.replace('\n','')))
			string = orig_string[1:-1]
			new_string = self.obfuscateString(string)
			info("\tObfuscated: (%s) => (%s...)" % (string.replace('\n',''), new_string.replace('\n','')[:80]))

			self.output = self.output.replace(orig_string, new_string)

	def obfuscateArrays(self):
		for m in re.finditer(r"\bArray\s*\(([^\)]+?)\)", self.output, flags=re.I|re.M):
			array = m.group(1)
			array = array.replace('\n', '').replace('\t', '')
			info("Array to obfuscate: Array(%s, ..., %s)" % (array[:40], array[-40:]))
			for_nums_allowed = [x for x in string.digits[:]]
			for_nums_allowed.extend(['a', 'b', 'c', 'd', 'e', 'f', ' ', '-', '_', '&', ','])
			if all(map(lambda x: x in for_nums_allowed, array)):
				# Array of numbers
				array = array.replace(' ', '').replace('_', '')
				try:
					info("\tLooks like array of numbers.")
					new_array = []
					nums = array.split(',')
					for num in nums:
						num = num.strip()
						new_array.append(self.obfuscateNumber(int(num)))

					obfuscated = 'Array(' + ','.join(new_array) + ')'
					info("\tObfuscated array: Array(%s, ..., %s)" % (','.join(new_array)[:40], ','.join(new_array)[-40:]))
					self.output = re.sub(r"Array\s*\(([^\)]+?)\)", obfuscated, self.output, flags=re.I | re.M)

				except ValueError as e:
					info("\tNOPE. This is not an array of numbers. Culprit: ('%s', context: '%s')" % (num, m.group(0)))
					continue
			else:
				info("\tThis doesn't seems to be array of numbers. Other types not supported at the moment.")


	def insertGarbage(self):
		if self.garbage_perc == 0.0:
			return

		lines = self.output.split('\n')
		garbages_num = int((self.garbage_perc / 100.0) * len(lines))
		new_lines = ['' for x in range(len(lines) + garbages_num)]
		garbage_lines = [random.randint(0, len(new_lines)-1) for x in range(garbages_num)]

		info('Appending %d garbage lines to the %d lines of input code %s' % \
			(garbages_num, len(lines), str(garbage_lines)))

		is_end = lambda x: ('End Sub' in x or 'End Function' in x)
		is_start = lambda x: ('Sub ' in x or 'Function ' in x) and ('(' in x or '()' in x)

		j = 0
		inside_func = False
		for i in range(len(new_lines)):
			if j >= len(lines): break
			line = lines[j]

			# TODO: THIS SIMPLY DOESNT WORK AT THE MOMENT.
			# if is_start(lines[j]) \
			# 	or (j > 0 and is_start(lines[j-1])) \
			# 	or (j > 1 and is_start(lines[j-2])) \
			# 	or (is_end(lines[j])) \
			# 	or (j < (len(lines) - 2) and is_end(lines[j+1])) \
			# 	or (j < (len(lines) - 3) and is_end(lines[j+2])):
			# 	inside_func = True

			# elif is_end(lines[j]) \
			# 	or (j > 0 and is_end(lines[j-1])) \
			# 	or (j > 1 and is_end(lines[j-2])) \
			# 	or (j > 2 and is_end(lines[j-3])) \
			# 	or (j > (len(lines) - 2) and is_start(lines[j+1])) \
			# 	or (j < (len(lines) - 3) and is_start(lines[j+2])) \
			# 	or (j < (len(lines) - 4) and is_start(lines[j+3])):
			# 	inside_func = False

			if i in garbage_lines:
				# Insert garbage
				comment = True 		# TODO: Switch this to False and improve inside_func code
				if inside_func:
					# Add comment or fake string initialization line.
					if random.randint(0, 2) == 0:
						comment = True
				else:
					comment = True

				varName = randomString(random.randint(4,12))
				varContents = self.obfuscateString(randomString(random.randint(10,30)))
				if comment:
					new_lines[i] = '\'Dim %(varName)s\n\'Set %(varName)s = %(varContents)s' % \
					{'varName' : varName, 'varContents' : varContents}
				else:
					new_lines[i] = 'Dim %(varName)s\nSet %(varName)s = %(varContents)s' % \
					{'varName' : varName, 'varContents' : varContents}
			else:
				new_lines[i] = lines[j]
				j += 1


		self.output = '\n'.join(new_lines)


def randomString(len):

	rnd = ''.join(random.choice(
		string.letters + string.digits
	) for _ in range(len))

	if rnd[0] in string.digits:
		rnd = random.choice(string.letters) + rnd

	return rnd

def parse_options(argv):
	global config
	args = None

	parser = argparse.ArgumentParser(
		prog = 'obfuscate.py',
		description = '''
	Attempts to obfuscate an input visual basic script in order to prevent
	curious eyes from reading over it.
''')

	group = parser.add_mutually_exclusive_group()
	group2 = parser.add_mutually_exclusive_group()
	parser.add_argument("input_file", help="Visual Basic script to be obfuscated.")
	parser.add_argument("-o", "--output", help="Output file. Default: stdout", default='')
	group2.add_argument("-g", "--garbage", help="Percent of garbage to append to the obfuscated code. Default: 12%%.", default=config['garbage_perc'], type=float)
	group2.add_argument("-G", "--no-garbage", dest="nogarbage", help="Don't append any garbage.", action='store_true')
	parser.add_argument("-m", "--min-var-len", dest='min_var_len', help="Minimum length of variable to include in name obfuscation. Too short value may break the original script. Default: 5.", default=config['min_var_length'], type=int)
	parser.add_argument("-r", "--reserved", action='append', help='Reserved word/name that should not be obfuscated (in case some name has to be in original script cause it may break it otherwise). Repeat the option for more words.')
	group.add_argument("-v", "--verbose", help="Verbose output.", action="store_true")
	group.add_argument("-q", "--quiet", help="No unnecessary output.", action="store_true")

	args = parser.parse_args()

	if args.verbose:
		config['verbose'] = True

	if not args:
		parser.print_help()

	if not os.path.isfile(args.input_file):
		err('Input file does not exist!')
		return False
	else:
		config['file'] = args.input_file

	if args.output:
		config['output'] = args.output

	if args.quiet:
		config['quiet'] = True

	if args.garbage <= 0.0 or args.garbage > 100.0:
		err("Garbage parameter must be in range (0, 100)!")
		return False
	else:
		config['garbage_perc'] = args.garbage

	if args.nogarbage:
		config['garbage_perc'] = 0.0

	if args.min_var_len < 0:
		err("Minimum var length must be greater than 0!")
	else:
		config['min_var_len'] = args.min_var_len

	if args.reserved:
		for r in args.reserved:
			config['custom_reserved_words'].append(r)

	return True

def main(argv):

	if not parse_options(argv):
		return False

	out('''
	Visual Basic script obfuscator for penetration testing usage.
	Mariusz B. / mgeeky, '17
''')

	ok('Input file:\t\t"%s"' % config['file'])
	if config['output']:
		ok('Output file:\t"%s"' % config['output'])
	else:
		ok('Output file:\tstdout.')

	contents = ''
	with open(config['file'], 'r') as f:
		contents = f.read()

	out('\n[.] Input file length: %d' % len(contents))

	obfuscator = ScriptObfuscator(
		config['custom_reserved_words'], \
		config['garbage_perc'], \
		config['min_var_length'])
	obfuscated = obfuscator.obfuscate(contents)

	if obfuscated:
		out('[.] Obfuscated file length: %d' % len(obfuscated))
		if not config['output']:
			out('\n' + '-' * 80)
			print obfuscated
			out('-' * 80)
		else:
			with open(config['output'], 'w') as f:
				f.write(obfuscated)
			out('[+] Obfuscated code has been written to: "%s"' % config['output'])
	else:
		return False


if __name__ == '__main__':
	main(sys.argv)
