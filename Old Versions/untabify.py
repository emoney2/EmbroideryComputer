# untabify.py
import sys

if len(sys.argv) != 2:
    print("Usage: python untabify.py <path-to-python-file>")
    sys.exit(1)

infile = sys.argv[1]
# Read with utf-8, replacing any invalid bytes
with open(infile, 'r', encoding='utf-8', errors='replace') as f:
    lines = f.readlines()

# Write back, converting each leading tab to 4 spaces
with open(infile, 'w', encoding='utf-8', newline='\n') as f:
    for line in lines:
        leading_tabs = len(line) - len(line.lstrip('\t'))
        new_line = '    ' * leading_tabs + line.lstrip('\t')
        f.write(new_line)
