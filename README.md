## VisualBasicObfuscator
Visual Basic Code universal Obfuscator intended to be used during penetration testing assignments.

```
$ ./obfuscate.py -h 
usage: obfuscate.py [-h] [-o OUTPUT] [-N | -g GARBAGE | -G] [-m MIN_VAR_LEN]
                    [-r RESERVED] [-v | -q]
                    input_file

Attempts to obfuscate an input visual basic script in order to prevent curious
eyes from reading over it.

positional arguments:
  input_file            Visual Basic script to be obfuscated.

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        Output file. Default: stdout
  -N, --normalize       Don't perform obfuscation, do only code normalization
                        (like long strings transformation).
  -g GARBAGE, --garbage GARBAGE
                        Percent of garbage to append to the obfuscated code.
                        Default: 12%.
  -G, --no-garbage      Don't append any garbage.
  -m MIN_VAR_LEN, --min-var-len MIN_VAR_LEN
                        Minimum length of variable to include in name
                        obfuscation. Too short value may break the original
                        script. Default: 5.
  -r RESERVED, --reserved RESERVED
                        Reserved word/name that should not be obfuscated (in
                        case some name has to be in original script cause it
                        may break it otherwise). Repeat the option for more
                        words.
  -v, --verbose         Verbose output.
  -q, --quiet           No unnecessary output.
```

---

## Disclaimer

There is much work to be done in this tool, so at the moment I'm not describing it thoroughly until it get's satisfiable level of functionality. 


## Obfuscation example

**Original form** (as comes from [here](https://gist.github.com/mgeeky/3c705560c5041ab20c62f41e917616e6)):

```

Const MY_CONSTANTS = "Some super const"
Variable As String
Dim AnotherVariable As String

Sub MyTestSub (ByVal arg1 As String)
    Query = "SELECT * FROM __InstanceModificationEvent WITHIN 60 " _
    & "WHERE TargetInstance ISA 'Win32_PerfFormattedData_PerfOS_System' " _
    & "AND TargetInstance.SystemUpTime >= 200 AND " _
    & "TargetInstance.SystemUpTime < 320"

    Dim arr As Variant
    arr = Array(1, 2, 3, 4, 5, 6)   ' This is an inline comment

    '
    ' This is a comment
    '
    Dim longString
    longString = "this is some very long string concatenated."
    longString = longString + "this is some very long string concatenated."
    longString = longString + "this is some very long string concatenated."
    longString = longString + "this is some very long string concatenated."
    longString = longString + "this is some very long string concatenated."
    longString = longString + "this is some very long string concatenated."

    Dim somecondition As Boolean
    somecondition = False
    If somecondition <> True Then
        End Sub
    End If

End Sub
```

---

**Obfuscated form**:

```
'Dim aVfvyu
'Set aVfvyu = Chr(Int("117"))&Chr(88)&Chr(Int("109"))&Chr(Int("&H62"))&"f"&Chr(-2940+2988)&"U"&"X"&Chr(-340+415)&Chr(Int("77"))&Chr(&H52)&Chr(Int("&H59"))&Chr(143330/1303)&Chr(&H65)&"y"
'Dim ISTfUwgvSFrf
'Set ISTfUwgvSFrf = "b"&Chr(Int("117"))&Chr(Int("67"))&"m"&"9"&Chr(2968-2892)&Chr(&H52)&Chr(&H56)&Chr(66)&Chr(&H4a)&"P"&Chr(Int("56"))&Chr(Int("&H42"))&Chr(115)&Chr(6602-6513)&Chr(88)&Chr(Int("&H75"))
Const pyMy8dwtX = Chr(Int("83"))&"o"&"m"&Chr(Int("101"))&Chr(38560/1205)&Chr(Int("&H73"))&Chr(Int("&H75"))&Chr(112)&"e"&Chr(114)&Chr(&H20)&Chr(99)&Chr(111)&Chr(&H6e)&Chr(Int("115"))&Chr(Int("&H74"))
HFaqG1lD6KKx As String
Dim Ano
K9FRiZbopz4n = Chr(&H53)&Chr(Int("&H45"))&Chr(&H4c)&Chr(69)&Chr(Int("67"))&Chr(Int("84"))&Chr(32)&Chr(&H2a)&Chr(32)&Chr(2509-2439)&Chr(82)&"O"&Chr(-100+177)&Chr(Int("32"))&"_"&Chr(-2119+2214)&Chr(Int("&H49"))&Chr(Int("110"))&Chr(Int("&H73"))&Chr(Int("&H74"))&Chr(&H61)&"n"&Chr(99)&"e"&Chr(Int("77"))&Chr(111)&Chr(Int("100"))&Chr(Int("&H69"))&"f"&Chr(Int("&H69"))&Chr(Int("99"))&Chr(Int("97"))&"t"&Chr(1680-1575)&Chr(Int("&H6f"))&Chr(&H6e)&Chr(Int("69"))&Chr(118)&Chr(Int("101"))&"n"&"t"&Chr(&H20)&"W"&Chr(73)&"T"&Chr(72)&Chr(Int("&H49"))&Chr(Int("78"))&Chr(Int("&H20"))&Chr(3113-3059)&Chr(48)&Chr(32)&Chr(-2805+2892)&Chr(Int("&H48"))&Chr(Int("69"))&"R"&Chr(Int("69"))&Chr(&H20)&Chr(&H54)&Chr(174891/1803)&Chr(114)&Chr(&H67)&"e"&Chr(Int("116"))&Chr(&H49)&Chr(Int("&H6e"))&Chr(Int("&H73"))&Chr(Int("&H74"))&Chr(8479-8382) _
& Chr(Int("&H6e"))
K9FRiZbopz4n = K9FRiZbopz4n + Chr(192753/1947)&Chr(101)&Chr(Int("32"))&Chr(73)&Chr(83)&Chr(65)&Chr(86016/2688)&"
& "t"
K9FRiZbopz4n = K9FRiZbopz4n + Chr(Int("&H65"))&Chr(&H6d)&Chr(Int("&H55"))&Chr(112)&"T"&Chr(Int("105"))&Chr(&H6d)&Chr(Int("101"))&Chr(&H20)&Chr(Int("&H3e"))&Chr(1833-1772)&" "&Chr(Int("&H32"))&Chr(&H30)&"0"&Chr(Int("&H20"))&Chr(Int("65"))&Chr(&H4e)&Chr(68)&Chr(Int("&H20"))&Chr(Int("84"))&"a"&"r"&Chr(Int("103"))&Chr(&H65)&Chr(Int("&H74"))&Chr(&H49)&Chr(Int("&H6e"))&Chr(9433-9318)&Chr(&H74)&"a"&Chr(110)&Chr(99)&Chr(&H65)&Chr(Int("&H2e"))&Chr(Int("83"))&Chr(&H79)&Chr(Int("115"))&Chr(116)&"e"&"m"&Chr(Int("&H55"))&Chr(&H70)&Chr(84)&Chr(105)&Chr(Int("109"))&Chr(&H65)&Chr(32)&Chr(7233-7173)&Chr(32)&Chr(51)&Chr(Int("50"))&"0"
Dim arr As Variant
arr = Array(8812-8811,1048/524,30/10,2601-2597,-1938+1943,6)
Dim JNmFJc8DgvM2
JNmFJc8DgvM2 = Chr(Int("116"))&Chr(Int("104"))&Chr(Int("105"))&"s"&Chr(32)&Chr(105)&Chr(Int("&H73"))&" "&Chr(6213-6098)&Chr(Int("111"))&Chr(&H6d)&Chr(2371-2270)&Chr(4910-4878)&Chr(118)&"e"&Chr(114)&Chr(121)&Chr(-32+64)&"l"&Chr(Int("&H6f"))&Chr(Int("&H6e"))&Chr(&H67)&Chr(&H20)&Chr(&H73)&Chr(Int("&H74"))&Chr(114)&Chr(Int("105"))&Chr(Int("&H6e"))&Chr(&H67)&" "&Chr(504-405)&Chr(8358-8247)&Chr(Int("110"))&Chr(&H63)&Chr(Int("97"))&Chr(116)&Chr(&H65)&Chr(110)&Chr(97)&Chr(Int("&H74"))&Chr(Int("&H65"))&Chr(Int("100"))&"."
'Dim bsCHuctI
'Set bsCHuctI = "s"&"U"&"t"&Chr(67)&Chr(&H71)&Chr(Int("66"))&"m"&Chr(Int("112"))&Chr(&H41)&Chr(Int("73"))&Chr(&H55)&Chr(Int("&H6b"))&"d"&"Y"&Chr(Int("&H4b"))&Chr(155344/2128)&Chr(&H59)&Chr(Int("&H61"))&Chr(Int("119"))&Chr(219300/2924)&Chr(65)&Chr(63291/1241)&Chr(Int("&H7a"))&Chr(Int("&H39"))&Chr(140679/1617)&Chr(Int("&H6e"))&Chr(&H6d)&Chr(99)&Chr(71)
JNmFJc8DgvM2 = JNmFJc8DgvM2 + Chr(Int("116"))&Chr(Int("104"))&Chr(Int("105"))&"s"&Chr(32)&Chr(105)&Chr(Int("&H73"))&" "&Chr(6213-6098)&Chr(Int("111"))&Chr(&H6d)&Chr(2371-2270)&Chr(4910-4878)&Chr(118)&"e"&Chr(114)&Chr(121)&Chr(-32+64)&"l"&Chr(Int("&H6f"))&Chr(Int("&H6e"))&Chr(&H67)&Chr(&H20)&Chr(&H73)&Chr(Int("&H74"))&Chr(114)&Chr(Int("105"))&Chr(Int("&H6e"))&Chr(&H67)&" "&Chr(504-405)&Chr(8358-8247)&Chr(Int("110"))&Chr(&H63)&Chr(Int("97"))&Chr(116)&Chr(&H65)&Chr(110)&Chr(97)&Chr(Int("&H74"))&Chr(Int("&H65"))&Chr(Int("100"))&"."
JNmFJc8DgvM2 = JNmFJc8DgvM2 + Chr(Int("116"))&Chr(Int("104"))&Chr(Int("105"))&"s"&Chr(32)&Chr(105)&Chr(Int("&H73"))&" "&Chr(6213-6098)&Chr(Int("111"))&Chr(&H6d)&Chr(2371-2270)&Chr(4910-4878)&Chr(118)&"e"&Chr(114)&Chr(121)&Chr(-32+64)&"l"&Chr(Int("&H6f"))&Chr(Int("&H6e"))&Chr(&H67)&Chr(&H20)&Chr(&H73)&Chr(Int("&H74"))&Chr(114)&Chr(Int("105"))&Chr(Int("&H6e"))&Chr(&H67)&" "&Chr(504-405)&Chr(8358-8247)&Chr(Int("110"))&Chr(&H63)&Chr(Int("97"))&Chr(116)&Chr(&H65)&Chr(110)&Chr(97)&Chr(Int("&H74"))&Chr(Int("&H65"))&Chr(Int("100"))&"."
JNmFJc8DgvM2 = JNmFJc8DgvM2 + Chr(Int("116"))&Chr(Int("104"))&Chr(Int("105"))&"s"&Chr(32)&Chr(105)&Chr(Int("&H73"))&" "&Chr(6213-6098)&Chr(Int("111"))&Chr(&H6d)&Chr(2371-2270)&Chr(4910-4878)&Chr(118)&"e"&Chr(114)&Chr(121)&Chr(-32+64)&"l"&Chr(Int("&H6f"))&Chr(Int("&H6e"))&Chr(&H67)&Chr(&H20)&Chr(&H73)&Chr(Int("&H74"))&Chr(114)&Chr(Int("105"))&Chr(Int("&H6e"))&Chr(&H67)&" "&Chr(504-405)&Chr(8358-8247)&Chr(Int("110"))&Chr(&H63)&Chr(Int("97"))&Chr(116)&Chr(&H65)&Chr(110)&Chr(97)&Chr(Int("&H74"))&Chr(Int("&H65"))&Chr(Int("100"))&"."
JNmFJc8DgvM2 = JNmFJc8DgvM2 + Chr(Int("116"))&Chr(Int("104"))&Chr(Int("105"))&"s"&Chr(32)&Chr(105)&Chr(Int("&H73"))&" "&Chr(6213-6098)&Chr(Int("111"))&Chr(&H6d)&Chr(2371-2270)&Chr(4910-4878)&Chr(118)&"e"&Chr(114)&Chr(121)&Chr(-32+64)&"l"&Chr(Int("&H6f"))&Chr(Int("&H6e"))&Chr(&H67)&Chr(&H20)&Chr(&H73)&Chr(Int("&H74"))&Chr(114)&Chr(Int("105"))&Chr(Int("&H6e"))&Chr(&H67)&" "&Chr(504-405)&Chr(8358-8247)&Chr(Int("110"))&Chr(&H63)&Chr(Int("97"))&Chr(116)&Chr(&H65)&Chr(110)&Chr(97)&Chr(Int("&H74"))&Chr(Int("&H65"))&Chr(Int("100"))&"."
JNmFJc8DgvM2 = JNmFJc8DgvM2 + Chr(Int("116"))&Chr(Int("104"))&Chr(Int("105"))&"s"&Chr(32)&Chr(105)&Chr(Int("&H73"))&" "&Chr(6213-6098)&Chr(Int("111"))&Chr(&H6d)&Chr(2371-2270)&Chr(4910-4878)&Chr(118)&"e"&Chr(114)&Chr(121)&Chr(-32+64)&"l"&Chr(Int("&H6f"))&Chr(Int("&H6e"))&Chr(&H67)&Chr(&H20)&Chr(&H73)&Chr(Int("&H74"))&Chr(114)&Chr(Int("105"))&Chr(Int("&H6e"))&Chr(&H67)&" "&Chr(504-405)&Chr(8358-8247)&Chr(Int("110"))&Chr(&H63)&Chr(Int("97"))&Chr(116)&Chr(&H65)&Chr(110)&Chr(97)&Chr(Int("&H74"))&Chr(Int("&H65"))&Chr(Int("100"))&"."
Dim GkSxDgf As Boolean
GkSxDgf = False
If GkSxDgf <> True Then
End Sub
End If
End Sub
```

---

### TODO:

- Improve obfuscation routines
- Add random garbage insertion
- Add strings encoding/decoding routine based on some simple transformations
- Add strings array based obfuscation
- Add `Eval` based code obfuscation routines
- Offer the user functionality of storing encryption/decryption key in document's metadata.

