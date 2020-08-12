def avg(*numbers):
    sum=0
    for n in numbers:
        sum=sum+n
    average=sum/len(numbers)
    return average
numbers=()
n=0
while n<5:
  num=int(input("numbers:"))
  a=avg(num)
  n=n+1
print(a)