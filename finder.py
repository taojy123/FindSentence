import re
import zipfile

filename = raw_input("Please enter the input file (default: a.docx):")

if not filename:
    filename = "a.docx"

if filename[-4:] == "docx":
    z = zipfile.ZipFile("a.docx")
    a = z.read("word/document.xml")
    z.close()
    a = re.sub(r"<.*?>", '', a) 
    a = a.decode("utf8").encode("gbk", "ignore")
elif filename[-4:] == ".txt":
    a = open(filename).read()


keyword = raw_input("Please enter the keyword:")


print "."

s = ""
sentences = a.strip().split('\xa1\xf9') # "\xe2\x80\xbb" for utf8
for sentence in sentences:
    sentence = sentence.strip()
    if keyword in sentence:
        s += sentence + "\n"

print ".."

open("result.txt", "w").write(s)

print "..."

raw_input("Finish, the result is in result.txt")