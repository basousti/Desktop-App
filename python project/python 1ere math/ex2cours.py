import matplotlib.pyplot as plt
import numpy as np
x=np.arange(0,5,0.01)
y=np.cos(x)**2
y1=np.cos(2*x)
y2=np.cos(x**2)
plt.plot(x,y,"-.",lw=2,color="pink",label="f(x)")
plt.plot(x,y1,lw=3,color ="black",label="g(x)")
plt.plot(x,y2,lw=4,color="yellow",label="h(x)")
plt.legend()
plt.show()
plt.close()