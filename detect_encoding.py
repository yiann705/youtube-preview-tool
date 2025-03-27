import chardet

with open('GG.csv', 'rb') as f:
    raw_data = f.read()
    result = chardet.detect(raw_data)
    print(result) 