# a = [1, 2, [3, 4]]
# print(len(a))
# def fun(a1, b):
#     return a1+b

# for median_age in range(0, 100, 10):
#     print(median_age)

# print("你好")

def sumList(listArr):
    if listArr == []:
        return None
    if len(listArr) == 1:
        return listArr[-1]
    else:
        e = listArr.pop()
        return e + sum(listArr)

listArr = [1, 2, 3, 4, 5, 6, 7]

result = sumList(listArr)


# listArr.pop()
print(result)
