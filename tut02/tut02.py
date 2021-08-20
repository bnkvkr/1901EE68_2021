def get_memory_score(list):
    count = 0
    temp = []
    for x in list:
        if x in temp:
            count = count + 1
        else:
            if(len(temp) == 5):
                temp.pop(0)
                temp.append(x)
            else:
                temp.append(x)
    return count


input_nums = [3, 4, 1, 6, 3, 3, 9, 0, 0, 0]
wrong = []
for x in input_nums:
    if(type(x) != int):
        wrong.append(x)

if(len(wrong) > 0):
    print("Please enter a valid input list")
    print("Invalid inputs detected:", wrong)
else:
    print("Score:", get_memory_score(input_nums))
