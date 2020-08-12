numbers = [1,3,5,7,10,9,6,12,15,20,25,30,24]
for n in filter(lambda n: n%3==0 or n%5==0,numbers):
    print (n)