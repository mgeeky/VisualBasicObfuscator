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
Public alreadyLaunched As Integer


Private Sub Malware()
    '
    ' ============================================
    '
    ' Enter here your malware code here.
    ' It will be started on auto open surely.
    '
    ' ============================================

    MsgBox ("Here comes the malware!")

    ' ============================================

End Sub


Private Sub Launch()
    If alreadyLaunched = True Then
        Exit Sub
    End If
    Malware
    SubstitutePage
    alreadyLaunched = True
End Sub

Private Sub SubstitutePage()
    '
    ' This routine will take the entire Document's contents,
    ' delete them and insert in their place contents defined in
    ' INSERT -> Quick Parts -> AutoText -> named as in `autoTextTemplateName`
    '
    Dim doc As Word.Document
    Dim firstPageRange As Range
    Dim rng As Range
    Dim autoTextTemplateName As String

    ' This is the name of the defined AutoText prepared in the document,
    ' to be inserted in place of previous contents.
    autoTextTemplateName = "RealDoc"

    Set firstPageRange = Word.ActiveDocument.Range
    firstPageRange.Select
    Selection.WholeStory
    Selection.Delete Unit:=wdCharacter, Count:=1

    Set doc = ActiveDocument
    Set rng = doc.Sections(1).Range
    doc.AttachedTemplate.AutoTextEntries(autoTextTemplateName).Insert rng, True
    doc.Save

End Sub

Sub AutoOpen()
    ' Becomes launched as first on MS Word
    Launch
End Sub

Sub Document_Open()
    ' Becomes launched as second, another try, on MS Word
    Launch
End Sub

Sub Auto_Open()
    ' Becomes launched as first on MS Excel
    Launch
End Sub

Sub Workbook_Open()
    ' Becomes launched as second, another try, on MS Excel
    Launch
End Sub
```

---

**Obfuscated form**:

```
Public JlN7SFfGa As Integer
Private Sub gYx6oac()
MsgBox (Chr(72)&Chr(&H65)&Chr(114)&Chr(101)&Chr(Int("32"))&Chr(99)&Chr(Int("&H6f"))&Chr(109)&Chr(101)&Chr(115)&Chr(6710-6678)&Chr(9995-9879)&Chr(104)&Chr(Int("101"))&Chr(Int("32"))&"m"&Chr(1849-1752)&Chr(Int("108"))&"w"&Chr(97)&Chr(114)&Chr(&H65)&Chr(&H21))
End Sub
Private Sub T3pFL()
If JlN7SFfGa = True Then
Exit Sub
End If
'Dim u3mUck
'Set u3mUck = "Z"&Chr(Int("83"))&Chr(71)&"8"&Chr(&H62)&"U"&Chr(&H6f)&Chr(Int("&H4f"))&"H"&"c"&Chr(91+6)&Chr(103)&Chr(114)&"A"&"y"&"d"&"O"&Chr(76)&Chr(57)&Chr(5862-5814)&Chr(56100/1100)&Chr(Int("&H6f"))&"f"&Chr(138376/2824)&"s"&Chr(Int("55"))&Chr(75)&Chr(&H44)
gYx6oac
V31A0Upk
JlN7SFfGa = True
End Sub
Private Sub V31A0Upk()
'Dim W6Fb
'Set W6Fb = Chr(Int("&H61"))&Chr(97)&"x"&Chr(Int("106"))&Chr(Int("&H6e"))&Chr(&H48)&Chr(Int("50"))&Chr(67)&Chr(Int("&H58"))&"N"&"l"&Chr(89)&Chr(Int("&H49"))&Chr(Int("65"))&Chr(&H30)&"8"&Chr(117)&"o"
Dim Yir4Xyg As Word.Document
Dim xIFlARxw9 As Range
'Dim ynGkjkU4wynR
'Set ynGkjkU4wynR = Chr(Int("&H77"))&Chr(Int("113"))&Chr(172050/1550)&Chr(Int("121"))&"5"&Chr(&H31)&Chr(69)&Chr(Int("&H30"))&Chr(Int("&H33"))&"o"&Chr(Int("&H45"))&Chr(4942-4824)&Chr(-2536+2652)
Dim j92Aq As Range
Dim N931oGuPbNa As String
N931oGuPbNa = Chr(Int("&H52"))&"e"&Chr(Int("&H61"))&Chr(906-798)&Chr(Int("&H44"))&Chr(Int("&H6f"))&"c"
Set xIFlARxw9 = Word.ActiveDocument.Range
xIFlARxw9.Select
Selection.WholeStory
Selection.Delete rVgJ2570=wdCharacter, rGG5pP3KL2=1
Set Yir4Xyg = ActiveDocument
Set j92Aq = Yir4Xyg.Sections(1).Range
Yir4Xyg.AttachedTemplate.AutoTextEntries(N931oGuPbNa).Insert j92Aq, True
Yir4Xyg.Save
'Dim h9RGN2pbYn
'Set h9RGN2pbYn = "J"&Chr(9055-8943)&Chr(77)&Chr(240198/2451)&Chr(82)&Chr(Int("&H63"))&Chr(Int("&H71"))&Chr(Int("&H53"))&Chr(114)&Chr(65)&Chr(114)&Chr(&H52)&"e"&Chr(112)&Chr(111)&Chr(Int("113"))&Chr(&H62)&"s"&Chr(117)&Chr(Int("90"))
End Sub
'Dim TkUaeM0Trg
'Set TkUaeM0Trg = Chr(Int("&H77"))&"8"&"q"&"3"&"c"&Chr(&H36)&Chr(Int("&H5a"))&Chr(Int("&H64"))&Chr(51)&Chr(Int("68"))&Chr(2446-2376)
Sub AutoOpen()
T3pFL
End Sub
Sub Document_Open()
T3pFL
End Sub
Sub Auto_Open()
T3pFL
End Sub
'Dim ipoqLK
'Set ipoqLK = Chr(&H6d)&Chr(Int("&H31"))&Chr(&H4d)&Chr(Int("109"))&Chr(57)&Chr(Int("&H72"))&Chr(5286-5198)&Chr(Int("57"))&Chr(Int("&H33"))&"Y"&Chr(Int("89"))&Chr(2400-2322)&Chr(&H52)&Chr(&H74)&"G"&"w"&Chr(49)&Chr(&H39)&Chr(68)&Chr(Int("52"))&Chr(Int("118"))
'Dim b3rTO
'Set b3rTO = Chr(67)&Chr(-1125+1173)&Chr(4944-4846)&Chr(Int("99"))&Chr(2448-2348)&Chr(84)&Chr(&H6a)&"v"&Chr(Int("&H6a"))&"r"&Chr(Int("100"))&Chr(Int("&H36"))&Chr(Int("&H39"))&"Z"&"f"
Sub Workbook_Open()
T3pFL
End Sub
'Dim tQEK5
'Set tQEK5 = Chr(Int("&H6f"))&Chr(51)&Chr(88)&Chr(2538-2459)&Chr(97)&"9"&Chr(Int("82"))&Chr(Int("53"))&Chr(Int("85"))&Chr(Int("&H31"))&Chr(Int("103"))&Chr(&H69)&Chr(&H4a)&Chr(Int("90"))&Chr(&H45)&Chr(-161+214)&Chr(224532/2673)&"b"
```

---

### TODO:

- Improve obfuscation routines
- Add random garbage insertion
- Add strings encoding/decoding routine based on some simple transformations
- Add strings array based obfuscation
- Add `Eval` based code obfuscation routines
- Offer the user functionality of storing encryption/decryption key in document's metadata.

