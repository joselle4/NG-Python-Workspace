def input_text(txt):
    print(txt)

while(1):
    try:
        input_text(input())
    except KeyboardInterrupt:
        print("Program Terminated")
        break
