def convert_number(arg):
    'use try/except clause to determine whether an arbitrary text string can be converted to a number'

    try:
        int(arg)
        print('String conversion successful. The number is ',arg)
    except ValueError:
        print('Cannot convert %s to a number.' % arg)


convert_number('6')
convert_number('blah')
convert = list(map(convert_number,['234','Joselle','8636','Abagat']))

#another way is to use an if statement to check if the value is a digit
def num_convert(arg):
    if not arg.isdigit():
        print('Cannot convert %s to a number.' % arg)
    else:
        print('String conversion successful. The number is ',arg)

convert = list(map(num_convert,['642','human','98','animals','this word']))
