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

	def __init__(self):
		self.input = ''
		self.output = ''

	def obfuscate(self, inp):
		self.input = inp
		self.output = inp

		# Step 1: Remove comments and indents
		self.output = re.sub(r"(?<!\"[^\"])'(.*)", "", self.output, flags=re.I)
		self.output = re.sub(r"\t| {2,}", "", self.output, flags=re.I)

		# Step 2: Remove empty lines
		self.output = '\n'.join(filter(lambda x: not re.match(r'^\s*$', x), self.output.split('\n')))

		# Step 3: Rename used variables
		self.randomizeVariablesAndFunctions()

		# Step 4: Explode string constants
		self.obfuscateStrings()

		return self.output

	def randomizeVariablesAndFunctions(self):
		for m in re.finditer(r"([ \t]*(?:Dim|Set|Sub)?\s*)(?<!\.)\b([^'\"\s\.]+)(\s*=)", self.output, flags = re.I|re.M):
			varName = randomString(random.randint(4,12))
			varToReplace = m.group(2)
			info("Variable name obfuscated: '%s' => '%s'" % (varToReplace, varName))
			self.output = self.output.replace(varToReplace, varName)

		for m in re.finditer(r"\s*(?:\w+)?(?:Sub|Function)\s+(\w+)\(\)", self.output, flags = re.I|re.M):
			varName = randomString(random.randint(4,12))
			varToReplace = m.group(1)
			if varToReplace in RESERVED_NAMES:
				continue
			info("Function name obfuscated: '%s' => '%s'" % (varToReplace, varName))
			self.output = self.output.replace(varToReplace, varName)

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
		new_string = ''
		for char in string:
			new_string += self.obfuscateChar(char) + '&'
		return new_string[:-1]

	def obfuscateStrings(self):
		for m in re.finditer(r"(\"[^\"]+\")", self.output, flags=re.I|re.M):
			orig_string = m.group(1)
			info("String to obfuscate: " + orig_string)
			string = orig_string[1:-1]
			new_string = self.obfuscateString(string)

			self.output = self.output.replace(orig_string, new_string)


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
	group.add_argument("-o", "--output", help="Output file. Default: stdout")
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

	return True

def main(argv):
	out('''
		Visual Basic script obfuscator for penetration testing usage.
	Mariusz B. / mgeeky, '17
''')

	if not parse_options(argv):
		return False

	ok('Input file:\t"%s"' % config['file'])
	if config['output']:
		ok('Output file:\t"%s"' % config['output'])
	else:
		ok('Output file:\tstdout.')

	contents = ''
	with open(config['file'], 'r') as f:
		contents = f.read()

	out('\n[.] Input file length: %d' % len(contents))

	obfuscator = ScriptObfuscator()
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