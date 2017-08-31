#!/usr/bin/python

import re
import os
import sys
import base64
import struct
import ctypes
import string
import random
import argparse

VERSION='0.1'

DEBUG = False
SPLIT = 80		# split long lines at this column
MAX_LINE_LENGTH = 1024 - 100

config = {
	'verbose' : False,
	'quiet' : False,
	'file' : '',
	'output' : '',
	'garbage_perc' : 12.0,
	'min_var_length' : 5,
	'custom_reserved_words': [],
	'normalize_only': False,
	'colors': True
}

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def out(x, col = ''):
	if not config['quiet']:
		if col:
			sys.stderr.write(col + x + bcolors.ENDC + '\n')
		else:
			sys.stderr.write(x + '\n')

def log(x, col = ''):
	if config['verbose']:
		out(x)

def err(x, col = ''):
	col2 = bcolors.BOLD + bcolors.FAIL if (config['colors'] and not col) else col
	out(col2 + '[!] ' + x + bcolors.ENDC)

def dbg(x, col = ''):
	if DEBUG:
		col2 = bcolors.HEADER if (config['colors'] and not col) else col
		log(col2 + "[DBG] " + x + bcolors.ENDC)

def info(x, col = ''):
	col2 = bcolors.OKBLUE if (config['colors'] and not col) else col
	log(col2 + '[?] ' + x + bcolors.ENDC)

def ok(x, col = ''):
	col2 = bcolors.BOLD + bcolors.OKGREEN if (config['colors'] and not col) else col
	out(col2 + '[+] ' + x + bcolors.ENDC)


class BitShuffleStringObfuscator:

	# Based on bits shuffling scheme devised by D. Knuth:
	#	(method described in D.Knuth's vol.4a chapter 7.1.3)
	OBFUSCATION_PARAMS = {
		'mask1': 0x00550055,
		'mask2': 0x0000cccc,
		'd1': 7,
		'd2': 14,
	}

	STRING_PADDING_CHAR = '~'

	DEOBFUSCATE_ROUTINE_NAME = 'MacroStringDeobfuscate'

	def __init__(self, obfuscateChar = None, obfuscateNumber = None):
		self.obfuscateChar = obfuscateChar
		self.obfuscateNum = obfuscateNumber

		if not obfuscateChar:
			self.obfuscateChar = BitShuffleStringObfuscator.obfuscateCharWrapper

		if not obfuscateNumber:
			self.obfuscateNum = BitShuffleStringObfuscator.obfuscateNumWrapper

	@staticmethod
	def obfuscateCharWrapper(x):
		return '"{}"'.format(x)

	@staticmethod
	def obfuscateNumWrapper(x):
		return '{}'.format(x)

	@staticmethod
	def composeDword(a, b, c, d):
		return (ord(a) << 24) | (ord(b) << 16) | (ord(c) << 8) | ord(d)

	@staticmethod
	def decomposeDword(num):
		a = (num & 0xff000000) >> 24
		b = (num & 0x00ff0000) >> 16
		c = (num & 0x0000ff00) >> 8
		d = (num & 0x000000ff)
		return (a,b,c,d)

	def obfuscateString(self, string, addDeobfName = True):
		chars = list(string)
		obfuscated = ''

		# Add null-byte padding for strings of length not divisible by 4
		if len(chars) % 4 != 0:
			chars.extend([BitShuffleStringObfuscator.STRING_PADDING_CHAR] * (4 - (len(chars) % 4)))

		for i in range(len(chars) / 4 + 1):
			fr = i * 4
			to = fr + 4
			if fr > len(chars) or to > len(chars): break
			dword0 = chars[fr:to][:]
			dword0.reverse()
			dword = BitShuffleStringObfuscator.composeDword(*dword0)
			obfuscatedDword = self.uintObfuscate(dword)
			obfuscated += struct.pack('<I', obfuscatedDword)

		if addDeobfName:
			return BitShuffleStringObfuscator.DEOBFUSCATE_ROUTINE_NAME + \
				'("%s")' % base64.b64encode(obfuscated)
		else:
			return base64.b64encode(obfuscated)

	def deobfuscateString(self, string):
		raw = list(base64.b64decode(string))
		out = ''

		for i in range(len(raw) / 4 + 1):
			fr = i * 4
			to = fr + 4
			if fr > len(raw) or to > len(raw): break
			dword0 = raw[fr:to][:]
			dword0.reverse()
			dword = BitShuffleStringObfuscator.composeDword(*dword0)
			rawDword = self.uintRestore(dword)
			out += struct.pack('<I', rawDword)

		# Remove null-byte padding
		while out.endswith(BitShuffleStringObfuscator.STRING_PADDING_CHAR):
			out = out[:-1]
		
		return out

	def uintObfuscate(self, num):
		mask1 = self.OBFUSCATION_PARAMS['mask1']
		mask2 = self.OBFUSCATION_PARAMS['mask2']
		d1 = self.OBFUSCATION_PARAMS['d1']
		d2 = self.OBFUSCATION_PARAMS['d2']

		t = ctypes.c_uint((num ^ (num >> d1)) & mask1).value
		u = ctypes.c_uint((num ^ t ^ (t << d1))).value
		t = ctypes.c_uint((u ^ (u >> d2)) & mask2).value
		out = ctypes.c_uint(u ^ t ^ (t << d2)).value
		return out

	def uintRestore(self, num):
		mask1 = self.OBFUSCATION_PARAMS['mask1']
		mask2 = self.OBFUSCATION_PARAMS['mask2']
		d1 = self.OBFUSCATION_PARAMS['d1']
		d2 = self.OBFUSCATION_PARAMS['d2']

		t = ctypes.c_uint((num ^ (num >> d2)) & mask2).value
		u = ctypes.c_uint(num ^ t ^ (t << d2)).value
		t = ctypes.c_uint((u ^ (u >> d1)) & mask1).value
		out = ctypes.c_uint(u ^ t ^ (t << d1)).value
		return out

	def getDeobfuscatorFuncName(self):
		return self.DEOBFUSCATE_ROUTINE_NAME

	def getDeobfuscatorCode(self):
		return '''
Public Function shrLong(ByVal Value As Long, ByVal Shift As Byte) As Long
	shrLong = Value
	If Shift > 0 Then
		If Value > 0 Then
			shrLong = Int(shrLong / (2 ^ Shift))
		Else
			If Shift > 31 Then
				shrLong = 0
			Else
				shrLong = shrLong And &H7FFFFFFF

				shrLong = Int(shrLong / (2 ^ Shift))

				shrLong = shrLong Or 2 ^ (31 - Shift)
			End If
		End If
	End If
End Function

Public Function shlLong(ByVal Value As Long, ByVal Shift As Byte) As Long
	shlLong = Value
	If Shift > 0 Then
		Dim i As Byte
		Dim m As Long
		For i = 1 To Shift
			' save 30th bit
			m = shlLong And &H40000000
			' clear 30th and 31st bits
			shlLong = (shlLong And &H3FFFFFFF) * 2
			If m <> 0 Then
				' set 31st bit
				shlLong = shlLong Or &H80000000
			End If
		Next i
	End If
End Function

Public Function DeobfuscateDword(ByVal num As Long) As Long
    Const mask1 As Long = %(mask1)d
    Const mask2 As Long = %(mask2)d
    Const d1 = 7
    Const d2 = 14
    Dim t As Long, u, out As Long
    
    ' Based on bits shuffling scheme devised by D. Knuth:
    '   (method described in D.Knuth's vol.4a chapter 7.1.3)
    t = (num Xor shrLong(num, d2)) And mask2
    u = num Xor t Xor shlLong(t, d2)
    t = (u Xor shrLong(u, d1)) And mask1
    out = (u Xor t Xor shlLong(t, d1))
    DeobfuscateDword = out
    
End Function

Public Function %(deobfuscateFunctionName)sHelper(ByRef inputString() As Byte) As String
    Dim i, fr, dword, raw As Long
    Dim a As String, b As String, c As String, d As String
    Dim deobf As String
    Dim adword() As String
    Dim a2, b2 As String

    deobf = ""

    For i = 0 To (UBound(inputString) / 4 + 1)
        fr = i * 4
        If fr > UBound(inputString) Then
            Exit For
        End If

        dword = 0
        dword = dword Or shlLong(inputString(fr + 3), 24)
        dword = dword Or shlLong(inputString(fr + 2), 16)
        dword = dword Or shlLong(inputString(fr + 1), 8)
        dword = dword Or inputString(fr + 0)
        
        raw = DeobfuscateDword(dword)
        
        a = Chr(shrLong((raw And &HFF000000), 24))
        b = Chr(shrLong((raw And 16711680), 16))
        c = Chr(shrLong((raw And 65280), 8))
        d = Chr(shrLong((raw And 255), 0))
        deobf = deobf + d + c + b + a
    Next i

    %(deobfuscateFunctionName)sHelper = deobf
End Function

Public Function %(deobfuscateFunctionName)s(inputString As String) As String

    Dim varByte1() As Byte, arrayByte2() As Byte, arrayByte3(255) As Byte
    Dim arrayLong4(63) As Long, arrayLong5(63) As Long
    Dim arrayLong6(63) As Long, bigNumber As Long
    Dim padding As Integer, iter As Long, index As Long, varLong3 As Long
    Dim deobf As String

    inputString = Replace(inputString, vbCr, vbNullString)
    inputString = Replace(inputString, vbLf, vbNullString)

    varLong3 = Len(inputString) Mod 4
    
    ' Check padding
    If InStrRev(inputString, "==") Then
        padding = 2
    ElseIf InStrRev(inputString, "" + "=") Then
        padding = 1
    End If

    ' Generate index arrays
    For varLong3 = 0 To 255
        Select Case varLong3
            Case 65 To 90
                arrayByte3(varLong3) = varLong3 - 65
            Case 97 To 122
                arrayByte3(varLong3) = varLong3 - 71
            Case 48 To 57
                arrayByte3(varLong3) = varLong3 + 4
            Case 43
                arrayByte3(varLong3) = 62
            Case 47
                arrayByte3(varLong3) = 63
        End Select
    Next varLong3

    ' Generate index translation arrays
    For varLong3 = 0 To 63
        arrayLong4(varLong3) = varLong3 * 64
        arrayLong5(varLong3) = varLong3 * 4096
        arrayLong6(varLong3) = varLong3 * 262144
    Next varLong3

    arrayByte2 = StrConv(inputString, vbFromUnicode)
    ReDim varByte1((((UBound(arrayByte2) + 1) \ 4) * 3) - 1)

    ' Convert big nums to quartets of raw chars
    For iter = 0 To UBound(arrayByte2) Step 4
        bigNumber = arrayLong6(arrayByte3(arrayByte2(iter))) + arrayLong5(arrayByte3(arrayByte2(iter + 1))) + arrayLong4(arrayByte3(arrayByte2(iter + 2))) + arrayByte3(arrayByte2(iter + 3))
        varLong3 = bigNumber And 16711680
        varByte1(index) = varLong3 \ 65536
        varLong3 = bigNumber And 65280
        varByte1(index + 1) = varLong3 \ 256
        varByte1(index + 2) = bigNumber And 255
        index = index + 3
    Next iter

    deobf = StrConv(varByte1, vbUnicode)
    If padding Then deobf = Left$(deobf, Len(deobf) - padding)
    %(deobfuscateFunctionName)s = %(deobfuscateFunctionName)sHelper(StrConv(deobf, vbFromUnicode))
    %(deobfuscateFunctionName)s = RemoveTrailingChars(%(deobfuscateFunctionName)s, "%(trailingCharacter)c")
End Function

Function RemoveTrailingChars(str As String, chars As String) As String
    Dim count As Long
    Dim parts() As String
    parts = Split(str, chars)
    count = UBound(parts, 1)
    If count <> 0 Then
        str = Left$(str, Len(str) - count)
    End If
    RemoveTrailingChars = str
End Function
''' % {
	'mask1': self.OBFUSCATION_PARAMS['mask1'], 
	'mask2': self.OBFUSCATION_PARAMS['mask2'], 
	'deobfuscateFunctionName' : BitShuffleStringObfuscator.DEOBFUSCATE_ROUTINE_NAME, 
	'trailingCharacter' : BitShuffleStringObfuscator.STRING_PADDING_CHAR
}


class ScriptObfuscator:

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

	COMMENT_REGEX = r"'((?!\").)*$"
	STRINGS_REGEX = r"(\"(?:[^\"\n]|\"\")+\")"

	FUNCTION_REGEX = r"(?:Public|Private|Protected|Friend)?\s*(?:Sub|Function)\s+(\w+)\s*\(.*\)"

	# var = "value" _\n& "value"
	LONG_LINES_REGEX = r"^\s*(?:(?:(\w+)\s*=(?:\s*\1\s*\+)?)|&)\s*\"([^\"]+)\"(?:\s+_)?"

	# Dim var ; Dim Var As Type ; Set Var = [...]
	VARIABLES_REGEX = r"^\s*(?:(?:\s*(\w+)\s*=(?!\"))|(?:Dim|Set|Const)\s+(\w+)\s*(?:As|=)?)|(?:^(\w+)\s+As\s+)"

	# Dim Var1, var2, va3 As Type
	VARIABLE_DECLARATIONS_REGEX = r"Dim\s+(?:(?:\s*,\s*)?(\w+)(?:\s*,\s*)?)+"

	FUNCTION_PARAMETERS_REGEX = r"(?:(?:Optional\s+)?\s*(?:ByRef|ByVal\s+)?(\w+)\s+As\s+\w+\s*,?\s*)"

	# Public Dim Var As Type ; Public Declare PtrSafe [...]
	GLOBALS_REGEX = r"\s*(?:(?:Public|Private|Protected)\s*(?:Dim|Set|Const)?\s+(\w+)\s*As)|(?:(?:Private|Protected|Public)\s+Declare\s+PtrSafe?\s*Function\s+(\w+)\s+)"


	def __init__(self, normalize_only = False, reserved_words = [], garbage_perc = 12.0, min_var_length = 5):
		self.input = ''
		self.output = ''

		self.bitShuffleObfuscator = BitShuffleStringObfuscator(ScriptObfuscator.obfuscateChar, ScriptObfuscator.obfuscateNumber)
		self.function_boundaries = []
		self.deobfuscatorAddedOnce = False
		self.avoidRemovingTheseComments = []

		self.normalize_only = normalize_only
		self.garbage_perc = garbage_perc
		self.min_var_length = min_var_length
		self.reserved_words = reserved_words

	def obfuscate(self, inp):
		self.input = inp
		self.output = inp

		# Normalization processes, avoid obfuscating.
		if self.normalize_only:
			self.output = self.mergeAndConcatLongLines(self.output)
			return self.output

		# Remove empty lines
		self.output = self.removeEmptyLines(self.output)

		# Merge long string lines and split them to concatenate.
		self.output = self.mergeAndConcatLongLines(self.output)

		# Explode string constants
		self.obfuscateStrings()

		# Remove comments
		self.removeComments()

		# Rename used variables
		self.randomizeVariablesAndFunctions()

		# Obfuscate arrays
		self.obfuscateArrays()

		# Insert garbage
		# TODO: Garbage insertion is flawed at the moment, resulting in inserting
		# 		junk lines that breaks line continuations (lines ending with '_').
		#self.insertGarbage()

		# Remove indents and multi-spaces.
		self.output = self.removeIndents(self.output)
		self.output = self.removeEmptyLines(self.output)

		return self.output

	def addDeobfuscator(self):
		if not self.deobfuscatorAddedOnce:
			deobfuscatorFunction = self.bitShuffleObfuscator.getDeobfuscatorCode()
			#deobfuscatorFunction = self.removeIndents(deobfuscatorFunction)
			deobfuscatorFunction = self.removeEmptyLines(deobfuscatorFunction)
			info("Appending bit shuffle string deobfuscation routines.")
			self.output += '\n%s' % deobfuscatorFunction
			self.deobfuscatorAddedOnce = True

	def removeEmptyLines(self, txt):
		return '\n'.join(filter(lambda x: not re.match(r'^\s*$', x), txt.split('\n')))

	def removeIndents(self, txt):
		if DEBUG:
			return txt
		txt = re.sub(r"\t| {2,}", "", txt, flags=re.I)
		lines = txt.split('\n')
		newLines = []
		# Have to be this way instead of finditer due to regex wrongly matching line breaks
		for i in range(len(lines)):
			line = lines[i]
			outLine = line
			for m in re.finditer(r"\s*(\+|\-|\/|\*|\=|\\|\,|\>|\<|\<\>|\^)\s*(?!\_)", line, flags=re.I):
				if '_' in line[m.span()[1]:]: continue
				outLine = outLine.replace(m.group(0), m.group(1))
			newLines.append(outLine)
		txt = '\n'.join(newLines)
		return txt

	def removeComments(self):
		lines = self.output.split('\n')
		for i in range(len(lines)):
			line = lines[i]
			m = re.search(r"('.*)", lines[i], flags=re.I)
			if not m: 
				continue

			# BUG: This loop does not detect comments added by `obfuscateString` to surround 
			#		'Declare PtrSafe Function' instructions. The loop cannot detect whether m.group(1)
			#		that is currently regex-matched comment has been marked with avoid-removing by being
			#		added to the `self.avoidRemovingTheseComments` list.
			
			# avoidComment = False
			# for avoid in self.avoidRemovingTheseComments:
			# 	if m.group(1) in avoid:
			# 		dbg("\tThis comment is marked to be avoided from removing: {{ %s }}" % m.group(1))
			# 		avoidComment = True
			# 		break

			# if avoidComment: continue

			dbg("\tChecking if comment (%s) is inside of a string: (%s)" % (m.group(1), lines[i]))
			n = re.search(r"(\"[^\"']*('[^\"]*)\")", lines[i], flags=re.I)

			if not n:
				dbg("\tNope, it's not.")
				o = re.match(r"^\s*'.*", m.group(1), flags=re.I)
				if o:
					dbg("\tSince the comment opens the line, the entire line gets wiped out.")
					lines[i] = ''
				else:
					lines[i] = line.replace(m.group(1), '')
				info("Found comment: (%s)" % m.group(1))
			else:
				dbg("\tYes it is. (group: (%s))" % n.group(1))

				if m.group(1) not in n.group(1) and m.group(1) not in lines[i]:
					info("Found comment: (%s)" % m.group(1))
					lines[i] = line.replace(m.group(1), '')

		self.output = '\n'.join(lines)

	def detectFunctionBoundaries(self):
		pos = 0
		del self.function_boundaries[:]
		funcStart = 0
		funcStop = 0
		counter = 0

		for m in re.finditer(ScriptObfuscator.FUNCTION_REGEX, self.output, flags=re.I|re.M):
			if pos >= len(self.output): break

			funcName = m.group(1)
			funcStart = m.span()[0] + (len(m.group(0)) - len(m.group(0).lstrip()))
			
			if counter > 0: 
				rev = self.output[pos:funcStart].rfind('End Sub')
				if rev == -1:
					rev = self.output[pos:funcStart].rfind('End Function')

				prev_end = pos + rev
				pos = prev_end
				self.function_boundaries[counter - 1]['funcStop'] = prev_end
				prev = self.function_boundaries[counter - 1]
				info("Function boundaries: (%s, from: %d, to: %d)" % (prev['funcName'], prev['funcStart'], prev['funcStop']))
			
			elem = {
				'funcName' : funcName,
				'funcStart' : funcStart,
				'funcStop' : -1
			}
			self.function_boundaries.append(elem)
			counter += 1

	def isInsideFunc(self, pos, offset = 0):
		for func in self.function_boundaries:
			if pos > (offset + func['funcStart']) and pos < (offset + func['funcStop']):
				return True
		return False

	def getFuncBoundaries(self, name):
		for func in self.function_boundaries:
			if name == func['funcName']:
				return func
		return None

	def findLongLines(self, txt, pos0, lineStart0 = 0, lineStop0 = 0):
		rex = ScriptObfuscator.LONG_LINES_REGEX
		pos = pos0
		varName = None
		longLine = None
		origLine = None
		lineStart = lineStart0
		lineStop = lineStop0
  
		dbg("Searching for long lines.", bcolors.OKBLUE)
		while pos < len(txt):
			dbg("Searching for long lines regex: POS = %d" % pos, bcolors.OKGREEN)
			m = re.search(rex, txt[pos:], flags=re.I|re.M)
			if not m: 
				break

			varName = m.group(1)
			longLine = m.group(2)

			endOfLine = pos + txt[pos:].find('\n')
			suffix = txt[pos:endOfLine + 1]
			
			lineStart = pos + m.span()[0]

			pos += m.span()[1]
			pos += len(suffix)		
			lineStop = pos - 1

			dbg("This is candidate for a long line (pos=%d, lineStart=%d, lineStop=%d, span: %s):\n{{ %s }}\n" % (pos, lineStart, lineStop, str(m.span()), longLine))
			assert lineStart < lineStop

			while True:
				if pos >= len(txt): break
				endOfLine = pos + txt[pos:].find('\n')
				line = txt[pos:endOfLine]
				dbg("\tDoes this line contains strings? (pos=%d, end=%d):\n\t{{ %s }}\n" % (pos,endOfLine,line))
				fault = False
				if endOfLine < pos:
					fault = True
					endOfLine, pos = pos, endOfLine
					line = txt[pos:]

				dbg("\tWe will find out by matching this line: pos=%d, endOfLine=%d,\n{{ %s }}" % (pos, endOfLine, line))
				n = re.match(rex, line, flags=re.I|re.M)
				if fault:
					pos += endOfLine + 1
				else:
					pos += endOfLine - pos + 1

				if not n: 
					dbg("\tOoops, not matching seemingly.")
					if len(longLine) > SPLIT:
						origLine = txt[lineStart:lineStop]
						dbg("Although this is helluva long line - %d bytes! Gotta break it up. (lineStart=%d, lineStop=%d, pos=%d):\n{{ %s }}\n" % (len(longLine), lineStart, lineStop, pos, longLine))
						return (varName, origLine, longLine, pos, lineStart, lineStop)
					break

				dbg("\tYes. It matches smoothly: (%s)" % n.group(2))
				longLine += n.group(2)
				lineStop = endOfLine
				
			assert lineStart < lineStop
			lineStop = pos

		origLine = txt[lineStart:lineStop]
		dbg("Found long line (lineStart=%d, lineStop=%d, pos=%d):\n{{ %s }}\n" % (lineStart, lineStop, pos, origLine))
		return (varName, origLine, longLine, pos, lineStart, lineStop)

	def mergeAndConcatLongLines(self, txt):
		replaces = []
		pos = 0
		lineStart = 0
		lineStop = 0
		
		while True:
			(varName, origLine, longLine, pos, lineStart, lineStop) = self.findLongLines(txt, pos, lineStart, lineStop)
			if not origLine or not longLine: break
			dbg("AFTER Searching for acompanying lines. Returned: varName = (%s), pos=%d, lineStart=%d, lineStop = %d\n\torigLine = {{ %s }}\n\tlongLine = {{ %s }}" % (varName, pos, lineStart, lineStop, origLine, longLine))
			newLine = ''
			for s in range(len(longLine) / SPLIT + 1):
				fr = s * SPLIT
				to = fr + SPLIT
				if newLine == '':
					newLine += '\n%s = "%s"\n' % (varName, longLine[fr:to])
				else:
					newLine += '%s = %s + "%s"\n' % (varName, varName, longLine[fr:to])

			dbg("Full long line joined altogether:\n{{ %s }}\n" % longLine)
			if len(longLine) < SPLIT:
				dbg("Too short to be replaced (%s)..." % longLine)
				continue
			else:
				dbg("This line will get merged:\n{{ %s }}\nINTO:\n{{ %s }}\n" % (origLine, newLine))
				replaces.append((origLine, newLine, longLine))
				if pos >= len(txt): break

		for (orig, new, longLine) in replaces:
			if DEBUG:
				dbg("Merging long line:\n\t{{ %s }}\n\t======>\n\t{{ %s }}\n\n" % (orig, new))
			else:
				info("Merging long string line (var: %s, len: %d): '%s...%s'" % (varName, len(longLine), longLine[:40], longLine[-40:]))
			txt = txt.replace(orig, new)

		return txt


	def randomizeVariablesAndFunctions(self):
		replacedAlready = {}

		def replaceVar(name, m):
			varName = randomString(random.randint(4,12))
			varToReplace = filter(lambda x: x and len(x)>0, m.groups())[0]

			if varToReplace in replacedAlready.keys(): return
			if len(varToReplace) < self.min_var_length: return
			if varToReplace in self.reserved_words: return

			info(name + " name obfuscated (context: \"%s\"): '%s' => '%s'" % (m.group(0).strip(), varToReplace, varName))
			replacedAlready[varToReplace] = varName

		# Variables
		for m in re.finditer(ScriptObfuscator.VARIABLES_REGEX, self.output, flags = re.I|re.M):
			replaceVar('Variable', m)

		for m in re.finditer(ScriptObfuscator.VARIABLE_DECLARATIONS_REGEX, self.output, flags = re.I|re.M):
			replaceVar('Variable', m)

		# Globals
		for m in re.finditer(ScriptObfuscator.GLOBALS_REGEX, self.output, flags = re.I|re.M):
			replaceVar('Global', m)


		for m in re.finditer(ScriptObfuscator.FUNCTION_REGEX, self.output, flags = re.I|re.M):
			pos = m.span()[0]
			funcName = m.group(1)

			for n in re.finditer(ScriptObfuscator.FUNCTION_PARAMETERS_REGEX, m.group(0), flags=re.I|re.M):
				pos2 = n.span()[0]
				varName = randomString(random.randint(4,12))
				varToReplace = n.group(1)

				if varToReplace in replacedAlready.keys(): continue
				replacedAlready[varToReplace] = varName
				info("Function: (%s): Adding variable obfuscation: (%s) => (%s)" % (funcName, varToReplace, varName))

		# Function names
		for m in re.finditer(ScriptObfuscator.FUNCTION_REGEX, self.output, flags = re.I|re.M):
			varName = randomString(random.randint(4,12))
			varToReplace = m.group(1)
			
			if len(varToReplace) < self.min_var_length: continue
			if varToReplace in self.reserved_words: continue
			if varToReplace in ScriptObfuscator.RESERVED_NAMES: continue
			if varToReplace in replacedAlready.keys(): continue
			
			info("Function name obfuscated (context: \"%s\"): '%s' => '%s'" % (m.group(0).strip(), varToReplace, varName))
			#self.output = re.sub(r"\b" + varToReplace + r"\b", varName, self.output, flags=re.I | re.M)
			replacedAlready[varToReplace] = varName

		for (varToReplace, varName) in replacedAlready.items():
			dbg("Obfuscate variable name: (%s) => (%s)" % (varToReplace, varName))
			self.output = re.sub(r"(?<!\.)(?:\b" + varToReplace + r"\b)|(?:\b" + varToReplace + r"\s*=\s*)", varName, self.output, flags=re.I | re.M)

		info("Randomized %d names in total." % len(replacedAlready))

	@staticmethod
	def obfuscateNumber(num):
		rnd1 = random.randint(0, 3333)
		num_coders = (
			lambda rnd1, num: '%d' % num,
			lambda rnd1, num: '%d-%d' % (rnd1+num, rnd1),
			lambda rnd1, num: '%d-%d' % (2*rnd1+num, 2*rnd1),
			lambda rnd1, num: '%d-%d' % (3*rnd1+num, 3*rnd1),
			lambda rnd1, num: '%d+%d' % (-rnd1, rnd1+num),
			lambda rnd1, num: '%d/%d' % (rnd1*num, rnd1),
		)

		out = random.choice(num_coders)(rnd1, num)
		if '/0' in out: 
			# Ooops, it cannot be happen!
			out = '%d' % num
		return out

	@staticmethod
	def obfuscateChar(char):
		char_coders = (
			lambda x: '"{}"'.format(x),	# 0
			lambda x: 'Chr(&H%x)' % ord(x),	# 1
			lambda x: 'Chr(%d)' % ord(x), # 2
			lambda x: 'Chr(%s)' % ScriptObfuscator.obfuscateNumber(ord(x)), # 3
			lambda x: 'Chr(Int("&H%x"))' % ord(x), # 4
			lambda x: 'Chr(Int("%d"))' % ord(x), # 5
		)
		return random.choice(char_coders)(char)

	def obfuscateString(self, string):
		return self.obfuscateStringBySubstitute(string)

	def obfuscateStringBySubstitute(self, string):
		if len(string) == 0: return ""
		new_string = ''
		delim = '&'
		if DEBUG: delim = ' & '

		i = 0
		while i < len(string):
			char = string[i]

			if len(new_string.split('\n')[-1]) + 128 > MAX_LINE_LENGTH:
				new_string = new_string[:-len(delim)] + ' _\n& '

			if i + 1 < len(string):
				if char == '"' and string[i+1] == '"':
					i += 2
					new_string += '""""' + delim
					continue
			if char == '"':
				# Fix improper quote escape
				new_string += '""""' + delim
				i += 1
				continue

			new_string += self.obfuscateChar(char) + delim
			i += 1

		new_string = new_string[:-len(delim)]
		if new_string.endswith(' _'): new_string = new_string[:-2]
		return new_string

	def obfuscateStrings(self, useBitShuffler = True):
		replaces = set()
		linesSurrounded = set()

		for m in re.finditer(ScriptObfuscator.STRINGS_REGEX, self.output, flags=re.I|re.M):
			orig_string = m.group(1)
			string = orig_string[1:-1]

			opening = max(m.span()[0]-500, 0)
			line_start = opening + self.output[opening:m.span()[0]].rfind('\n') + 1
			line_stop = m.span()[1] + self.output[m.span()[1]:].find('\n') + 1
			line = self.output[line_start:line_stop]

			if 'declare ' in line.lower() \
				and 'lib ' in line.lower() \
				and 'function ' in line.lower():
				# Syntax error while obfuscating pointer names and libs
				
				if line in linesSurrounded: continue
				# BUG: Surrounding 'Declare PtrSafe Function' with comments is messing 
				#		up `removeComments` procedure that gets called right after `obfuscateStrings`.
				#		The avoided-comments list approach is not working by now correctly.

				# if self.garbage_perc > 0:
				# 	varName = randomString(random.randint(8,20))
				# 	varName2 = randomString(random.randint(8,20))
				# 	junk = self.obfuscateString(randomString(random.randint(40,50)))
				# 	junk2 = self.obfuscateString(randomString(random.randint(40,50)))
				# 	garbage = '\'Dim %(varName)s\n\'Set %(varName)s = %(varContents)s\n' % \
				# 	{'varName' : varName, 'varContents' : junk}
				# 	garbage2 = '\'Dim %(varName)s\n\'Set %(varName)s = %(varContents)s\n' % \
				# 	{'varName' : varName2, 'varContents' : junk2}

				# 	info("Surrounding Pointer declaration with obfuscated junk")
				# 	dbg("\tJunk to surround:\nREPLACE this:\t{{ %s }}\nWITH this:\t{{ %s }}" % (line, garbage + line + garbage2))
				# 	replaces.add((line, garbage + line + garbage2))
				# 	linesSurrounded.add(line)

				# 	self.avoidRemovingTheseComments.extend([x for x in garbage.split('\n') if x])
				# 	self.avoidRemovingTheseComments.extend([x for x in garbage2.split('\n') if x])
				continue
			elif line.lstrip().startswith("'"): continue

			exceptionallyAvoidBitShuffler = False

			if 'const ' in line.lower():
				# Const are not to be anyhow obfuscated.
				continue

			if BitShuffleStringObfuscator.STRING_PADDING_CHAR in string:
				info("\tPadding character: (%s) has been detected in input string. Have to avoid Bit Shuffle string encoder." % BitShuffleStringObfuscator.STRING_PADDING_CHAR)
				exceptionallyAvoidBitShuffler = True

			info("String to obfuscate (context: {{%s}}, len: %d): {{%s}}" % (m.group(0).strip(), len(orig_string), string))
			
			if exceptionallyAvoidBitShuffler or not useBitShuffler or len(string) <= 5:
				new_string = self.obfuscateStringBySubstitute(string)
			else:
				new_string = '"%s"' % self.bitShuffleObfuscator.obfuscateString(string, False)
				if len(new_string) <= 40:
					new_string = self.obfuscateStringBySubstitute(new_string[1:-1])
				new_string = '%s(%s)' % (self.bitShuffleObfuscator.getDeobfuscatorFuncName(), new_string)

			info("OBFUSCATED:\n\t%s\n\t{{ %s }}\n\t=====>\n\t{{ %s }}\n\t%s\n" % ('^' * 60, string, new_string, '^' * 60))

			replaces.add((orig_string, new_string))

		for (orig, new) in replaces:
			dbg("Replacing:\n\t{{ %s }}\n\t=====>\n\t{{ %s }}\n" % (orig, new))
			self.output = self.output.replace(orig, new) 

		self.addDeobfuscator()

	def obfuscateArrays(self):
		for m in re.finditer(r"\bArray\s*\(([^\)]+?)\)", self.output, flags=re.I|re.M):
			orig_array = m.group(1)
			array = orig_array
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
						new_array.append(ScriptObfuscator.obfuscateNumber(int(num)))

					obfuscated = 'Array(' + ','.join(new_array) + ')'
					info("\tObfuscated array: Array(%s, ..., %s)" % (','.join(new_array)[:40], ','.join(new_array)[-40:]))
					self.output = re.sub(r"Array\s*\(" + orig_array + "\)", obfuscated, self.output, flags=re.I | re.M)

				except ValueError as e:
					info("\tNOPE. This is not an array of numbers. Culprit: ('%s', context: '%s')" % (num, m.group(0)))
					continue
			else:
				info("\tThis doesn't seems to be array of numbers. Other types not supported at the moment.")


	def insertGarbage(self):
		if self.garbage_perc == 0.0:
			return

		self.detectFunctionBoundaries()

		lines = self.output.split('\n')
		inside_func = False
		garbages_num = int((self.garbage_perc / 100.0) * len(lines))
		new_lines = ['' for x in range(len(lines) + garbages_num)]
		garbage_lines = [random.randint(0, len(new_lines)-1) for x in range(garbages_num)]

		info('Appending %d garbage lines to the %d lines of input code %s' % \
			(garbages_num, len(lines), str(garbage_lines)))

		j = 0
		pos = 0
		offset = 0

		for i in range(len(new_lines)):
			if j >= len(lines): break

			inside_func = self.isInsideFunc(pos, offset)

			if i in garbage_lines:
				if i > 0 and (\
					(' _' in new_lines[i - 1].rstrip()[-10:]) or \
					(' _' in new_lines[i - 2].rstrip()[-10:]) or \
					('&' in lines[j].lstrip()[:5]) \
				):
					# Avoid inserting garbage that could break line continuations
					continue

				comment = True
				if inside_func:
					comment = False

				varName = randomString(random.randint(4,12))
				varContents = self.obfuscateString(randomString(random.randint(10,30)))
				garbage = ''
				if comment:
					garbage = '\'Dim %(varName)s As String\n\'%(varName)s = %(varContents)s' % \
					{'varName' : varName, 'varContents' : varContents}
				else:
					garbage = 'Dim %(varName)s As String\n%(varName)s = %(varContents)s' % \
					{'varName' : varName, 'varContents' : varContents}

				# TODO:
				new_lines[i] = garbage
				offset += 2
			else:
				new_lines[i] = lines[j]
				pos += len(lines[j]) + 1
				j += 1

		self.output = '\n'.join(new_lines)


def classifyFileAndExtractContents(contents):
	# Simple decision tree
	htmlOrHtaFile = (
		'<html', '</html>', '<script', '</script>'
	)

	allMarkers = True
	for a in htmlOrHtaFile:
		if a not in contents.lower():
			allMarkers = False
			break

	if allMarkers:
		ok("File has been classified as HTML/HTA or containing <script> to be extracted.")
		m = re.search(r"\<script[^\>]+>(.*)</script>", contents, flags=re.I|re.M|re.S)
		if m:
			return m.group(1)
		else:
			err("Although it was not possible to extract script contents. Proceeding with raw")
			return contents

	return contents

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
	group2.add_argument("-N", "--normalize", dest="normalize_only", help="Don't perform obfuscation, do only code normalization (like long strings transformation).", action='store_true')
	group2.add_argument("-g", "--garbage", help="Percent of garbage to append to the obfuscated code. Default: 12%%.", default=config['garbage_perc'], type=float)
	group2.add_argument("-G", "--no-garbage", dest="nogarbage", help="Don't append any garbage.", action='store_true')
	group2.add_argument("-C", "--no-colors", dest="nocolors", help="Dont use colors.", action='store_true')
	parser.add_argument("-m", "--min-var-len", dest='min_var_len', help="Minimum length of variable to include in name obfuscation. Too short value may break the original script. Default: 5.", default=config['min_var_length'], type=int)
	parser.add_argument("-r", "--reserved", action='append', help='Reserved word/name that should not be obfuscated (in case some name has to be in original script cause it may break it otherwise). Repeat the option for more words.')
	group.add_argument("-v", "--verbose", help="Verbose output.", action="store_true")
	group.add_argument("-d", "--debug", help="Debug output.", action="store_true")
	group.add_argument("-q", "--quiet", help="No unnecessary output.", action="store_true")

	args = parser.parse_args()

	if args.verbose:
		config['verbose'] = True

	if args.debug:
		global DEBUG
		config['verbose'] = True
		DEBUG = True

	if not args:
		parser.print_help()

	if not os.path.isfile(args.input_file):
		err('Input file does not exist!')
		return False
	else:
		config['file'] = args.input_file

	if args.output:
		config['output'] = args.output

	if args.nocolors:
		config['colors'] = False

	if args.normalize_only:
		config['normalize_only'] = args.normalize_only

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
        ver: %(versionNum)s
''' % {'versionNum' : VERSION })

	ok('Input file: "%s"' % config['file'])
	if config['output']:
		ok('Output file: "%s"' % config['output'])
	else:
		ok('Output file:\tstdout.')

	contents = ''
	with open(config['file'], 'r') as f:
		contents = f.read()

	contents = classifyFileAndExtractContents(contents)

	ok('Input file length: %d' % len(contents))

	obfuscator = ScriptObfuscator(
		config['normalize_only'], \
		config['custom_reserved_words'], \
		config['garbage_perc'], \
		config['min_var_length'])
	obfuscated = obfuscator.obfuscate(contents)

	if obfuscated:
		ok('Obfuscated file length: %d' % len(obfuscated))
		if not config['output']:
			out('\n\n')
			ok('-' * 80 )

			print obfuscated

			ok('-' * 80)
			out('\n')
		else:
			with open(config['output'], 'w') as f:
				f.write(obfuscated)
			ok('Obfuscated code has been written to: "%s"\n' % config['output'])
	else:
		return False

if __name__ == '__main__':
	main(sys.argv)
