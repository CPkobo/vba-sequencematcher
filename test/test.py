from difflib import SequenceMatcher

seq = SequenceMatcher()
seg1 = "lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur"
seg2 = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur."
seq = SequenceMatcher("", seg1, seg2)
print("---RATIO---")
print(seq.ratio())
print("---QUICK_RATIO---")
print(seq.quick_ratio())
print("---REAL_QUICK_RATIO---")
print(seq.real_quick_ratio())
