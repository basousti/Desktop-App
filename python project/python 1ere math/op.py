import matplotlib.pyplot as plt

from math import sin , pi ,cos 
x=[]
y=[]
y1=[]
a=pi
pas=pi/200
for k in range (0,201):
    x.append(cos(a))
    y.append(sin(a))
    y1.append(-sin(a))
    a+=pas

plt.axis('equal')
plt.plot(x,y,x,y1)
plt.show()
plt.close()