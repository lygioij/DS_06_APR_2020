dict={"make":"audi",
      "model":"A4",
      "year":"2018",
      "color":"Green"}
for k in sorted(dict.values(),key=lambda k:k.lower()):
    print (k)
