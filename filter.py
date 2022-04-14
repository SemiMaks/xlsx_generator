from main import new_d


def total():
    print(new_d)
    for i in new_d:
        del new_d[0::6]
        print(new_d)


print(new_d)
total()
