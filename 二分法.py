import numpy as np
import easygui as eg


def fun(x):
    return np.exp(x)+x -4
box = eg.multenterbox('请输入取值范围','',['a','b'])
a = float(box[0])
b = float(box[1])
#a = float(1)
#b = float(2)
x = a +(b-a)/2
print('a','\t\t','b','\t\t','x','\t\t','b-a','\t\t','±')
print(a,'\t\t',b,'\t\t',x,'\t\t',b-a,'\t\t','-' if fun(a)*fun(x)< 0 else '+')
flag = fun(x)
ans = flag*fun(a)
num = 1 


while 1 :
    x= a +(b-a)/2
    flag = fun(x)
    ans = flag*fun(a)
    if ans < 0:
        b = x 
    else :
        a = x 
    now_x = a +(b-a)/2
    print(a,'\t\t',b,'\t\t',now_x,'\t\t',b-a,'\t\t','-' if fun(a)*fun(x)< 0 else '+', '         ' ,num)
    num += 1 

