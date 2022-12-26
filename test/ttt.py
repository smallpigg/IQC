rownum = [4, 3, 2, 1]
k = 0
for i in range(1, 4):
    k = k + 2
    for j in range(2, 7):
        rownum.append(k+7)
        k = k + 1

# print(rownum)


A1 = ['sdfs']
A2 = ['sdfsfs']
b = A1+A2

# print(b)

i = 2
str1 = '00'+str(i)
# print(str1)

for i in range(1, 7):
    str1 = '00' + str(i)
    # print(str1)

IQC_001_path = 'C:\\Users\\Zz\PycharmProjects\\IQC\\template\\IQC-001.docx'
# print(IQC_001_path)
# IQC_path = IQC_001_path.replace("001", str1, 1)
# print(IQC_path)

for i in range(1, 7):
    if i == 3:
        str1 = '00' + str(i)
        IQC_001_path
        IQC_path = IQC_001_path.replace("001", str1, 1)
        print(IQC_path)
        print("hello")
    else:
        print("hello111")