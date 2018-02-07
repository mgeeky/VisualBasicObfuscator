## VisualBasicObfuscator
Visual Basic Code universal Obfuscator intended to be used during penetration testing assignments.
To be used mainly to avoid AV and mail filters detections as well as Blue Teams inspection tasks.

Here is a sample **VirusTotal.com** scan against malicious Word document with embedded [RobustPentestMacro](https://github.com/mgeeky/RobustPentestMacro) geared up with _Empire_-generated powershell stager:

No obfuscation | Obfuscated 
--- | ---
[15 / 58](https://www.virustotal.com/#/file/d45af91e7a46cedc0aaa68a79eea48f19f47cdf2202e8347e61c178d987e2dcd/detection) | [9 / 58](https://www.virustotal.com/#/file/ea2e812c62543946f9f175b4183db2555d69673307bb046138b41e7fa9f63b91/detection)

- In case of _Non-obfuscated_ sample, the AV software has flagged the file with certain malicious-macro statement.
- As of obfuscated sample we've been left with _obfuscated_ flag.

There is still a huge area of improvement in testing which obfuscation techniques trigger what patterns, and work towards reducing such detection rate.


---

### FEATURES

- Able to extract HTML/HTA/<script> contents
- Able to obfuscate arrays of numbers and characters
- Obfuscating strings via Bit Shuffling and base64 encoding (_as described in D.Knuth's vol.4a chapter 7.1.3_). This method produces smaller in size results (approx. 66% smaller resulting scripts)
- Merging long concatenated lines into variables appendings to avoid maximum number of continuing lines (24)
- Junk insertion, smart enough to avoid breaking syntax outside of routines
- Sensitive to quote escapes within strings, detecting consecutive lines concatenation
- Variables, Globals, Constants, function names, function parameters names randomization
- Comments, indents and blank lines removal,



```
usage: obfuscate.py [-h] [-o OUTPUT] [-N | -g GARBAGE | -G | -C]
                    [-m MIN_VAR_LEN] [-r RESERVED] [-v | -q]
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
  -C, --no-colors       Dont use colors.
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


## Obfuscation example

**Original form** 

```
Const MY_CONSTANTS = "Some super const"
Dim AnotherVariable As String

Sub MyTestSub(ByVal arg1 As String)
    ' The below real-world used query contains nested quotes/apostrophes
    ' That could confuse `removeComments` routine
    Query = "SELECT * FROM __InstanceModificationEvent WITHIN 60 " _
    & "WHERE TargetInstance ISA 'Win32_PerfFormattedData_PerfOS_System' " _
    & "AND TargetInstance.SystemUpTime >= 200 AND " _
    & "TargetInstance.SystemUpTime < 320"

    Dim arr As Variant
    arr = Array(1, 2, 3, 4, 5, 6)   ' This is an inline comment

    Dim var1, var2
    var1 = "Testing Short string"
    var2 = "short"

    '
    ' This is a comment
    '
    Dim longString
    longString = "1. this is some very long string concatenated."
    longString = longString + "2. this is some very long string concatenated."
    longString = longString + "3. this is some very long string concatenated."
    longString = longString + "4. this is some very long string concatenated."

    longString = longString + "5. this is some very long string concatenated." _
    & "6. this is some very long string concatenated." _
    & "7. this is some very long string concatenated." _
    & "8. this is some very long string concatenated." _
    & "9. this is some very long string concatenated."

    Dim somecondition As Boolean
    somecondition = False
    If somecondition <> False Then
        Exit Sub
    End If

    MsgBox ("Test1(Constant): " & MY_CONSTANTS)
    MsgBox ("Test2(Query): " & Query)
    MsgBox ("Test3(var1 + var2): " & var1 & var2)
    MsgBox ("Test4(longString): " & longString)
    MsgBox ("Test5(Array's contents): " & Join(arr, ","))

End Sub

Sub test()
    MyTestSub (0)
End Sub
```

---

**Obfuscated form**:

Invocation - we have to specify that the macro is named `test` and therefore this Sub name shall not be mangled:

```
bash $ ./obfuscate.py -r test -v demo.vbs
```

Output (mind that _Constants_ contents cannot be anyhow obfuscated, as they have to be known at compile time):

```
Const MY_CONSTANTS="Some super const"
Dim H5AyJRiT As String
Sub LdzqEiHkpovt(ByVal arg1 As String)
K8fqa=onOrpZJSZL("Aq0oV8IwZBFohAVdSO6+TJjg8124pPFV0K9zUbqn0FD4q9xWwLr1VWiukEpQooxWoBBRCQ==") _
& onOrpZJSZL("EuqkUgCucEii6dFXmODzXbik8VU4hYBYaI3RSep3F0yIoVRf2qdRXaC8cF2g4sBVMLVhXSrsnVOC9+9aoK/cFw==") _
& onOrpZJSZL("EKhhEYinVF5gsvxV4rD0V6ptPFGgvXNeoPMiXbCp8hGyU1YDoBQQQg==") _
& onOrpZJSZL("iKdUXmCy/FXisPRXqm08UaC9c16g8yJdsKnyEahC1go=")
Dim arr As Variant
arr=Array(3256/3256,293-291,-1964+1967,-2979+2983,1678-1673,342-336)
Dim var1,var2
var1=onOrpZJSZL(Chr(Int("&H69"))&Chr(&H4f)&Chr(Int("&H50"))&Chr(Int("&H51"))&"X"&Chr(55)&Chr(Int("&H69"))&Chr(Int("111"))&Chr(Int("56"))&Chr(Int("120"))&Chr(&H47)&Chr(&H61)&Chr(&H36)&Chr(-209+289)&Chr(Int("120"))&Chr(Int("&H53"))&Chr(113)&Chr(Int("&H4f"))&Chr(76)&"Q"&Chr(Int("84"))&Chr(114)&"q"&Chr(&H74)&Chr(88)&Chr(70)&Chr(-2196+2285)&Chr(Int("&H3d")))
var2=Chr(Int("115"))&Chr(Int("104"))&Chr(111)&"r"&Chr(1900-1784)
Dim oAQ96qYI As String
oAQ96qYI=Chr(&H4c)&Chr(56)&Chr(Int("&H79"))&Chr(112)&Chr(111)&Chr(76)&Chr(Int("48"))&Chr(Int("&H50"))&Chr(78645/749)&Chr(&H45)&Chr(&H49)&Chr(110)&Chr(Int("78"))&Chr(Int("97"))&Chr(2431-2328)&Chr(&H52)&Chr(-1842+1955)&Chr(3538-3458)&"7"&Chr(Int("104"))&Chr(121)
Dim ChtxGWGe As String
ChtxGWGe=Chr(1513-1429)&Chr(80)&Chr(Int("69"))&Chr(121)&Chr(Int("122"))&Chr(&H41)&Chr(Int("&H73"))&Chr(Int("&H31"))&Chr(Int("&H76"))&"l"&"k"&Chr(Int("&H33"))&Chr(Int("67"))&Chr(&H42)&Chr(51)&Chr(&H78)&Chr(&H37)&Chr(&H71)&Chr(878-766)&Chr(2660-2571)&Chr(Int("71"))&Chr(Int("&H58"))&Chr(77)&Chr(113)&Chr(78276/1186)
Dim E1X9VVBEe
E1X9VVBEe=onOrpZJSZL("sEBxR7ih0higdXdQsqvyEernUFv4iNxVquLwTLqtXFboid1Uoq1wXKKvUlw=")
E1X9VVBEe=E1X9VVBEe+onOrpZJSZL("skBRR7ih0higdXdQsqvyEernUFv4iNxVquLwTLqtXFboid1Uoq1wXKKvUlw=")
E1X9VVBEe=E1X9VVBEe+onOrpZJSZL("skBxR7ih0higdXdQsqvyEernUFv4iNxVquLwTLqtXFboid1Uoq1wXKKvUlw=")
E1X9VVBEe=E1X9VVBEe+onOrpZJSZL("sEJRR7ih0higdXdQsqvyEernUFv4iNxVquLwTLqtXFboid1Uoq1wXKKvUlw=")
E1X9VVBEe=E1X9VVBEe+onOrpZJSZL("sEJxR7ih0higdXdQsqvyEernUFv4iNxVquLwTLqtXFboid1Uoq1wXKKvUlw=") _
& onOrpZJSZL("skJRR7ih0higdXdQsqvyEernUFv4iNxVquLwTLqtXFboid1Uoq1wXKKvUlw=") _
& onOrpZJSZL("skJxR7ih0higdXdQsqvyEernUFv4iNxVquLwTLqtXFboid1Uoq1wXKKvUlw=") _
& onOrpZJSZL("sEBTR7ih0higdXdQsqvyEernUFv4iNxVquLwTLqtXFboid1Uoq1wXKKvUlw=") _
& onOrpZJSZL("sEBzR7ih0higdXdQsqvyEernUFv4iNxVquLwTLqtXFboid1Uoq1wXKKvUlw=")
Dim EKyef8cVxSrr As Boolean
Dim v9yzH As String
v9yzH="t"&Chr(Int("&H62"))&Chr(Int("57"))&"r"&Chr(&H57)&Chr(Int("&H61"))&Chr(Int("&H4c"))&"W"&Chr(Int("&H6b"))&Chr(109)&"x"&Chr(108)&Chr(Int("50"))&Chr(115)&Chr(&H6e)&"P"&Chr(&H36)&Chr(80)&Chr(86)&Chr(Int("108"))&Chr(Int("&H74"))
EKyef8cVxSrr=False
If EKyef8cVxSrr<>False Then
Exit Sub
End If
Dim P0HBOB As String
P0HBOB=Chr(-1358+1446)&Chr(-3235+3316)&Chr(27010/365)&"7"&Chr(Int("&H6c"))&Chr(75)&"R"&Chr(6462-6396)&Chr(Int("&H57"))&Chr(86)&Chr(Int("106"))&Chr(99)&Chr(49)&Chr(&H56)&Chr(52496/772)&Chr(&H64)&"s"&Chr(Int("&H66"))&Chr(8944-8832)&Chr(Int("&H63"))&Chr(Int("104"))&Chr(72)&Chr(&H43)&Chr(Int("107"))&Chr(Int("&H4b"))&Chr(-1162+1278)&Chr(3073-2984)&"e"
MsgBox ("Test1(Constant): " & MY_CONSTANTS)
MsgBox (onOrpZJSZL(Chr(105)&Chr(79)&Chr(75840/948)&Chr(Int("&H51"))&Chr(Int("&H58"))&"z"&Chr(-344+420)&"E"&Chr(Int("48"))&"E"&Chr(55)&"g"&Chr(2905-2789)&Chr(-294+412)&Chr(&H6b)&Chr(Int("89"))) & K8fqa)
MsgBox (onOrpZJSZL(Chr(&H69)&Chr(79)&Chr(80)&Chr(&H51)&Chr(895-807)&Chr(6451-6396)&Chr(&H71)&"M"&Chr(Int("&H63"))&Chr(&H45)&Chr(114)&Chr(-1738+1843)&Chr(&H4e)&Chr(Int("&H56"))&"Q"&Chr(&H43)&Chr(283938/2558)&Chr(Int("78"))&"D"&Chr(86)&Chr(85)&"b"&Chr(Int("&H6f"))&Chr(Int("66"))&"W"&"A"&Chr(111)&Chr(61)) & var1 & var2)
MsgBox (onOrpZJSZL(Chr(105)&"O"&Chr(Int("80"))&Chr(2870-2789)&Chr(&H58)&Chr(2709-2662)&Chr(Int("&H43"))&Chr(79)&Chr(Int("88"))&Chr(Int("69"))&Chr(196601/2209)&Chr(Int("&H71"))&Chr(&H34)&Chr(Int("&H39"))&"N"&Chr(3900/39)&Chr(Int("&H75"))&Chr(113)&Chr(-1542+1591)&"c"&Chr(Int("86"))&Chr(189417/1839)&Chr(&H3d)&"=") & E1X9VVBEe)

' Cut for brevity. One can see the rest of this output in 'demo-obfuscated.vbs'
' [...]
```

---

### TODO:

- Improve obfuscation routines
- ~~Add strings encoding/decoding routine based on some simple transformations~~
- Add strings array based obfuscation
- Add polymorphic code obfuscation with given key to put into document's metadata
- Add `Eval` based code obfuscation routines
- Offer the user functionality of storing encryption/decryption key in document's metadata.
- Reinforce variable names randomization to have var names randomized each time within function-block 
- Implement several techniques for avoiding detections implemented in [Revoke-Obfuscation](https://www.blackhat.com/docs/us-17/thursday/us-17-Bohannon-Revoke-Obfuscation-PowerShell-Obfuscation-Detection-And%20Evasion-Using-Science-wp.pdf), mainly for fun and for avoidance of generic PUA (_Potentially Unwatend Application_) detection algorithms in modern Anti-Virus softwares


---

### KNOWN BUGS:

- HTA script encoding when Javascript was used, should not be possible. Also - more tests are needed whether VBS script encoding within HTA actually works all the time.
- There is a bug within `removeComments` being called after `obfuscateString` that has added comments to surround `Declare PtrSafe Function` instructions. Such dynamically added junk-comments should be marked not to be removed, or get added after calling `removeComments` instead.
- Inserting junk that is not a comment doesn't work properly (something about function boundaries detection)
    (also junk insertion breaks the syntax, so it had been **disabled** temporarily until the bug gets fixed).
- ~~There is a bug with function parameters mangling that introduces various syntax errors and breaks~~
- If there is a string within quotes placed in the comment - it will wrongly get obfuscated breaking the syntax
