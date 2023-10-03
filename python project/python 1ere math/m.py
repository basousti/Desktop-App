import matplotlib.pyplot as plt
from math import cos, pi 
x = [] 
y = [] 
pas = 2 * pi / 100 
abscisse = 0 
for k in range(0, 101): 
 x.append(abscisse) 
 y.append(cos(abscisse)) 
 abscisse += pas 
plt.plot(x, y)
plt.show()
plt.close()

#ou bien :
#from matplotlib.pylab import plot, arange, cos, show
#X = arange(-5,3,0.01) 
#Y = cos(X) 
#plot(X,Y) 
#show()
#close()