import glob
import docx


user_search = input("Enter an interest you are looking for:")


for file in glob.glob('wordfolder/*.docx'):
    doc = docx.Document(file)
    for i in file:
        result = [p.text for p in doc.paragraphs]
        if user_search.upper() in str(result).upper():
            print(file)
            print("match!")
            break
        else:
            print("No match!")
            break

