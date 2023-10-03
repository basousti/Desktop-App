import matplotlib.pyplot as plt
from math import*
a=-2
b=2
X=[]
Y=[]
pas=0.01
abscisse =a
for k in range (0,201):
    X.append(abscisse)
    Y.append(1/abscisse)
    abscisse+=pas
plt.axis([-1,1,-10,10])    
plt.title("etude de fonction ln")
plt.plot(X,Y,color="yellow") #plt.scatter(X,Y)->nuage des points 
plt.show()
plt.close()