#!/usr/bin/python

import re
import os
import sys
import string
import random
import argparse

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

config = {
	'verbose' : False,
	'quiet' : False,
	'file' : '',
	'output' : '',
	'garbage_perc' : 12.0,
	'min_var_length' : 5,
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

	def __init__(self, garbage_perc, min_var_length):
		self.input = ''
		self.output = ''
		self.garbage_perc = garbage_perc
		self.min_var_length = min_var_length

	def obfuscate(self, inp):
		self.input = inp
		self.output = inp

		# Step 1: Remove comments
		self.output = re.sub(r"(?<!\"[^\"])'(.*)", "", self.output, flags=re.I)

		# Step 2: Rename used variables
		self.randomizeVariablesAndFunctions()

		# Step 3: Explode string constants
		self.obfuscateStrings()

		# Step 4: Insert garbage
		self.insertGarbage()

		# Step 5: Remove empty lines
		self.output = '\n'.join(filter(lambda x: not re.match(r'^\s*$', x), self.output.split('\n')))
		
		# Step 6: Remove indents and multi-spaces.
		self.output = re.sub(r"\t| {2,}", "", self.output, flags=re.I)

		return self.output

	def randomizeVariablesAndFunctions(self):
		# Variables
		for m in re.finditer(r"(?:(?:Dim|Set|Const)\s+(\w+)\s*(?:As|=)?)|(?:^(\w+)\s+As\s+)", self.output, flags = re.I|re.M):
			varName = randomString(random.randint(4,12))
			varToReplace = filter(lambda x: x and len(x)>0, m.groups())[0]

			if len(varToReplace) < self.min_var_length: continue
			info("Variable name obfuscated (context: \"%s\"): '%s' => '%s'" % (m.group(0).strip(), varToReplace, varName))
			self.output = re.sub(r"\b" + varToReplace + r"\b", varName, self.output, flags=re.I | re.M)

		# Globals
		for m in re.finditer(r"\s*(?:(?:Public|Private|Protected)\s*(?:Dim|Set|Const)?\s+(\w+)\s*As)|(?:(?:Private|Protected|Public)\s+Declare\s+PtrSafe?\s*Function\s+(\w+)\s+)", self.output, flags = re.I|re.M):
			varName = randomString(random.randint(4,12))
			varToReplace = filter(lambda x: x and len(x)>0, m.groups())[0]

			if len(varToReplace) < self.min_var_length: continue
			info("Global name obfuscated (context: \"%s\"): '%s' => '%s'" % (m.group(0).strip(), varToReplace, varName))
			self.output = re.sub(r"\b" + varToReplace + r"\b", varName, self.output, flags=re.I | re.M)			

		# Functions
		for m in re.finditer(r"\s*(?:Public|Private|Protected|Friend)?\s*(?:Sub|Function)\s+(\w+)\s*\(", self.output, flags = re.I|re.M):
			varName = randomString(random.randint(4,12))
			varToReplace = m.group(1)
			if len(varToReplace) < self.min_var_length: continue
			if varToReplace in RESERVED_NAMES:
				continue
			info("Function name obfuscated (context: \"%s\"): '%s' => '%s'" % (m.group(0).strip(), varToReplace, varName))
			self.output = re.sub(r"\b" + varToReplace + r"\b", varName, self.output, flags=re.I | re.M)

	def obfuscateNumber(self, num):
		rnd1 = random.randint(0, 3333)
		num_coders = (
			lambda rnd1, num: '%d' % num,
			lambda rnd1, num: '%d-%d' % (rnd1+num, rnd1),
			lambda rnd1, num: '%d-%d' % (2*rnd1+num, 2*rnd1),
			lambda rnd1, num: '%d-%d' % (3*rnd1+num, 3*rnd1),
			lambda rnd1, num: '%d+%d' % (num-rnd1, rnd1),
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
	parser.add_argument("input_file", help="Visual Basic script to be obfuscated.")
	parser.add_argument("-o", "--output", help="Output file. Default: stdout")
	parser.add_argument("-g", "--garbage", help="Percent of garbage to append to the obfuscated code. Default: 12%%. 0 to disable.", default=config['garbage_perc'], type=float)
	parser.add_argument("-m", "--min-var-len", dest='min_var_len', help="Minimum length of variable to include in name obfuscation. Too short value may break the original script. Default: 5.", default=config['min_var_length'], type=int)
	group.add_argument("-v", "--verbose", help="Verbose output.", action="store_true")
	group.add_argument("-q", "--quiet", help="No output.", action="store_true")

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

	if args.garbage < 0.0 or args.garbage > 100.0:
		err("Garbage parameter must be in range (0, 100)!")
		return False
	else:
		config['garbage_perc'] = args.garbage

	if args.min_var_len < 0:
		err("Minimum var length must be greater than 0!")
	else:
		config['min_var_len'] = args.min_var_len

	return True

def main(argv):
	out('''
		Visual Basic script obfuscator for penetration testing usage.
	Mariusz B. / mgeeky, '17
''')

	if not parse_options(argv):
		return False

	ok('Input file:\t\t"%s"' % config['file'])
	if config['output']:
		ok('Output file:\t"%s"' % config['output'])
	else:
		ok('Output file:\tstdout.')

	contents = ''
	with open(config['file'], 'r') as f:
		contents = f.read()

	out('\n[.] Input file length: %d' % len(contents))

	obfuscator = ScriptObfuscator(config['garbage_perc'], config['min_var_length'])
	obfuscated = obfuscator.obfuscate(contents)

	if obfuscated:
		out('[.] Obfuscated file length: %d\n' % len(obfuscated))
		out('-' * 60)
		print obfuscated
		out('-' * 60)
	else:
		return False


if __name__ == '__main__':
	main(sys.argv)
