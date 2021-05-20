# VBA-SequenceMatcher

## Abstarct
Text diff scripts in VBA (Word), partially ported from Python difflib.SequenceMathcer.

Implemented methods:
- set_seq1
- set_seq2
- set_seq2
- find_longest_match
- get_matching_blocks
- sort_blocks *
- get_opcode
- calculate_ratio
- ratio
- quick_ratio
- real_quick_ratio

\* represents original method of vba version.
**NOTE that junk function has not benn implemented yet**

This repository is under the license of PSFL.

## Usage
"SequenceMather.cls" is the entity of SequenceMathcer. You can create a instance in any modules.
NOTE that junk function has not benn implemented yet

```vba
Dim seq As SequenceMatcher
Dim seg1 As String, seg2 As String

Set seq = New SequenceMatcher
seg1 = "lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur"
seg2 = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur."

Call seq.set_seqs(seg1, seg2)

Debug.Print "---RATIO---"
Debug.Print seq.ratio
Debug.Print "---QUICK_RATIO---"
Debug.Print seq.quick_ratio
Debug.Print "---REAL_QUICK_RATIO---"
Debug.Print seq.real_quick_ratio
```

This code results in:
```
---RATIO---
 0.989505231380463 
---QUICK_RATIO---
 0.989505231380463 
---REAL_QUICK_RATIO---
 0.998500764369965 
```

and native python goes:
```
---RATIO---
0.9895052473763118
---QUICK_RATIO---
0.9895052473763118
---REAL_QUICK_RATIO---
0.9985007496251874
```

## Applying Opcode

Another bas module "WordSeqApplyer" has a sub procedure named "apply_opcode".
It can rewrite the text in document with revision according to the opcode.
For more details, please see the file "test.docm", and run "main".

