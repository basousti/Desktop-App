import matplotlib.pyplot as plt
import numpy as np
x=np.arange(-3,3,0.01)
y=x**2
x1=np.arange(-3,3,0.01)
y1=1-x**2
plt.plot(x,y,color='black')
plt.plot(x1,y1)
#plt.axis ('equal')
plt.show()
plt.close()