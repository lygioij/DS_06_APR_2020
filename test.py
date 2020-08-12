name=input("enter name:")
phone=int(input("phone number:"))
phone=[phone]
directory={name:phone}
while True:
    name = input("enter name:")
    phone = int(input("phone number:"))
    if(name == 'end'):
        break
    if(name in directory):
        directory[name].append(phone)
    else:
        phone = [phone]
        directory[name]=phone

for name,phone in directory.items():
    print(name,':',phone)