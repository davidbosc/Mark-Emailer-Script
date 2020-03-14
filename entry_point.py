print("Select a mode to execute this script:")
print("---------------------------------------")
print("\t eml - for email (Outlook)")
print("\t txt - for txt files")
print("---------------------------------------")
print("\t q - to cancel")

validInput = False

while (not validInput):
    mode = input()
    if mode == "q":
        exit()
    elif mode == "eml":
        exec(open('send_grades_email.py').read())
        validInput = True
    elif mode == "txt":
        exec(open('send_grades_txt_docs.py').read())
        validInput = True
    else :
        print("Please enter valid input")