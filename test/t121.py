list_string = ['产品包装']
string_set = set(['材料', '产品包装', '单证资料', '规格型号', '合格证明'])

result = all([word in string_set for word in list_string])

print(result)