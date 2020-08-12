strings=[' this','finer ',' THIS ','AGree ']
for n in sorted(strings,key=lambda n:n.strip().lower()):
    print(n)