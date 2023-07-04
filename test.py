list = []

for i in range(0,3):
    list_x = []
    for j in range(0,5):
        #print(j)
        list_x.append(j)
        #print(list_x)
    list.append(list_x)

list2 = ["red","blue"]

print(list2)
list2.remove("red")
print(list2)

def dai_xie(weight,height,age):
    return 67 + 13.73 * weight + 5 * height -6.9 * age

print(dai_xie(61,185,25))


