import matplotlib.pyplot as plt
import numpy as np
x=np.arange(0,5,0.01)
y= np.cos(x)**2
y1= np.cos(2*x)
y2= np.cos(x**2)
plt.plot(x,y,color="pink",label="no1")
plt.plot(x,y1,color="black",label="no2")
plt.plot(x,y2,color="red",lw=5,linestyle="-.",label="no3")
plt.xlabel("abscisse")
plt.ylabel("ordonnees")
plt.legend()
plt.show()
plt.close()