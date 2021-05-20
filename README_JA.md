# VBA-SequenceMatcher

## 概要
Pythonのdifflib.SequenceMathcerの、VBA （Word）部分移植版です。

以下のメソッドを実装しています。
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

\* は独自の補助メソッドです。

**junk関連は未実装なのでご注意ください。**

このリポジトリはPSFLのライセンスとしています。

## 使い方
"SequenceMather.cls" がSequenceMathcerの実態です。任意のモジュールでSequenceMathcerをインスタンス化します。


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

上記のコードを実行すると、イミディエイトウィンドウに次のように表示されます。
```
---RATIO---
 0.989505231380463 
---QUICK_RATIO---
 0.989505231380463 
---REAL_QUICK_RATIO---
 0.998500764369965 
```

同じ文をオリジナルのPythonで実行した結果です。
```
---RATIO---
0.9895052473763118
---QUICK_RATIO---
0.9895052473763118
---REAL_QUICK_RATIO---
0.9985007496251874
```

## opcodeの適用

もう一つのbasモジュール"WordSeqApplyer"には、"apply_opcode"というサブプロシージャがあります。
これを呼び出すことで、Wordファイル上に履歴を残しながら差分箇所を反映することができます。
詳細については、"test.docm"を開き、"main"を実行してみてください。

