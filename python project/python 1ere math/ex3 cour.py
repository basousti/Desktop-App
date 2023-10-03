import matplotlib.pyplot as plt
import numpy as np
from math import*
x=[0,pi/6,pi/4,pi/3,pi/2,2*pi/3,3*pi/4,5*pi/6,pi,7*pi/6,5*pi/4,4*pi/3,3*pi/2,5*pi/3,7*pi/4,11*pi/6,0]
x1=np.cos(x)
y=np.sin(x)
plt.plot(x1,y)
plt.axis('equal')
plt.show()
plt.close()
