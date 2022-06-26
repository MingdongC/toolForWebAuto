# -*- coding: utf-8 -*-
import matplotlib.pyplot as plt

name_list = ['插入时间', '查找时间', '删除时间']
num_list = [1.103, 0.025, 0.034]
num_list1 = [2.336, 2.550, 2.522]

x = list(range(len(num_list)))
total_width, n = 0.8, 2
width = total_width / n

plt.bar(x, num_list, width=width, label='大小平衡树', fc='b')
for i in range(len(x)):
    x[i] = x[i] + width
plt.bar(x, num_list1, width=width, label='二叉排序树', tick_label=name_list, fc='r')
plt.legend()
plt.show()