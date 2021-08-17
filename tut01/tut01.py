

def meraki_helper(n):
    prev = n % 10
    n = n//10

    while(n > 0):
        next = n % 10
        if(abs(next - prev) != 1):
            return 0
        n = n//10
        prev = next
    return 1


count = 0
input = [12, 14, 56, 1]
for i in range(0, len(input)):
    ele = input[i]
    if (meraki_helper(ele) == 1):
        count = count + 1
        print("Yes -", ele, "is a Meraki number")
    else:
        print("No -", ele, "is not a Meraki number")

print("the input list contains", count,
      "meraki and", len(input)-count, "non meraki numbers.")
