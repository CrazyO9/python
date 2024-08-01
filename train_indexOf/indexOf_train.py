longStr = "its aa pencil"
shortStr = "a pen"

def strpos(withinStr,findStr):
    for idx in range(len(withinStr)):
        for i in range(len(findStr)):
            if withinStr[idx+i] == findStr[i]:
                if i == len(findStr)-1:
                    return idx
            else:
                break
idx = strpos(longStr,shortStr)
print(f"'{shortStr}' start at {idx}")
