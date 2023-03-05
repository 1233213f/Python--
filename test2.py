import sympy as sp
from sympy import *

fm, t0, tm, r, d, t = sp.symbols('fm t0 tm r d t')
f1 = fm
f2 = ((t + t0)/(tm + t0))**r
f31 = d/(d + r)
f321 = r/(d + r)
f322 = ((t + t0)/(tm + t0))**(r + 1)
f32 = (f321 * f322)
n = -(r + d)/(r + 1)
f3 = (f31 + f32)**n
fuc = f1 * f2 * f3
fuc_x = sp.diff(fuc, t)
# print("对x求导结果 : ", fuc)

def hu(fm, t0, tm, r, d):
    x = fm*r*((t + t0)/(t0 + tm))**r*((t + t0)/(t0 + tm))**(r + 1)*(-d - r)*(d/(d + r) + r*((t + t0)/(t0 + tm))**(r + 1)/(d + r))**((-d - r)/(r + 1))/((d + r)*(t + t0)*(d/(d + r) + r*((t + t0)/(t0 + tm))**(r + 1)/(d + r))) + fm*r*((t + t0)/(t0 + tm))**r*(d/(d + r) + r*((t + t0)/(t0 + tm))**(r + 1)/(d + r))**((-d - r)/(r + 1))/(t + t0)
    re = solve(x, t)
    n = len(re)
    v = [0 for _ in range(n)]
    for i in range(0, n):
        v[i] = fm * ((re[i] + t0) / (t0 + tm)) ** r * (d / (d + r) + r * ((re[i] + t0) / (t0 + tm)) ** (r + 1) / (d + r)) ** ((-d - r) / (r + 1))
        if v[i] == max(v):
            max_y = v[i]
            max_x = re[i]
    x1 = fm * ((t + t0) / (t0 + tm)) ** r * (d / (d + r) + r * ((t + t0) / (t0 + tm)) ** (r + 1) / (d + r)) ** ((-d - r) / (r + 1)) - max_y * 0.5
    re1 = solve(x1, t)
    bfz = max_y * 0.5
    # print(re1)
    # n1 = len(re1)
    # v1 = [0 for _ in range(n1)]
    # for i in range(0, n1):
    #     v1[i] = fm * ((re1[i] + t0) / (t0 + tm)) ** r * (d / (d + r) + r * ((re1[i] + t0) / (t0 + tm)) ** (r + 1) / (d + r)) ** ((-d - r) / (r + 1))
        # print(v1[i])

    fh = [0 for _ in range(5)]
    fh[0] = max_x
    fh[1] = max_y
    fh[2] = re1[0]
    fh[3] = re1[1]
    fh[4] = bfz
    return fh

pt = [0 for _ in range(10)]
for i in range(0, 10):
    pt[i] = [0 for _ in range(10)]
pt[0][0] = "峰值对应的x"
pt[0][1] = "峰值"
pt[0][2] = "左解x1"
pt[0][3] = "右解x2"

# print(pt)
# hu(1,1,1,1,1)
print(hu(0.08033,5.68373,0.67762,10.53091,2.98718))
# def hu(fm, t0, tm, r, d):