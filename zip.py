a=[1,2]
b=[3,4]
c=[5,6]
var=[a,b,c]
rez=zip(*var[::-1])
print(i for i in rez)
for i in rez:print(i)



