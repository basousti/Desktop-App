import matplotlib.pyplot as plt
import numpy as np
x=np.arange(-3,3,0.01)
y=x**2
y1=1-x**2
plt.plot(x,y ,x,y1)
plt.show()
plt.close()