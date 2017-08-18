
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

		# Step 1: Remove comments
		self.output = re.sub(r"?<!\"[^\"])'(.+)", self.input, flags=re.I)



def randomString(len):
	return ''.join(random.choice(
		string.ascii_uppercase 
		+ string.ascii_lowercase
		+ string.digits
	) for _ in range(len))

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
	with open(config['input'], 'r') as f:
		contents = f.read()

	out('\n[.] Input file length: %d' % len(contents))

	obfuscator = ScriptObfuscator()
	obfuscated = obfuscator.obfuscate(contents)

	if obfuscated:
		out('\n[.] Obfuscated file length: %d' % len(obfuscated))
	else:
		return False


if __name__ == '__main__':
	main(sys.argv)