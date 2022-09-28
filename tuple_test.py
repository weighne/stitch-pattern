tuple1 = (10, 10, 20)
tuple2 = (20, 20, 10)

print(tuple1)
print(tuple2)

tuple3 = (0, 0, 0)

for x in range(len(tuple1)):
    tuple3[x] += tuple1[x]

result = tuple(map(lambda x, y: x - y, tuple1, tuple2))

print(tuple3)
print(result)
