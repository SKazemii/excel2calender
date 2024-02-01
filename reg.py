import re

input_string = ""

# Define the regex pattern
pattern = r'\b(\d+)-(\d+)(?:\s*([a-zA-Z]+)?)'

# # Use re.findall to extract all matching numbers
# start, end , txt= re.findall(pattern, input_string)[0]
txt= re.findall(pattern, input_string)

# print(start, end, txt)

pattern = re.compile(r'(\d+)-(\d+)([a-zA-Z]+)')
match = pattern.match(input_string)

print(txt)
