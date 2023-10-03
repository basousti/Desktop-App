import matplotlib.pyplot as plt
from math import*
a=1
b=20
X=[]
Y=[]
pas=(b-a)/200
abscisse =a
for k in range (0,201):
    X.append(abscisse)
    Y.append(log(abscisse))
    abscisse+=pas
plt.title("etude de fonction ln")
plt.plot(X,Y,color="yellow") #plt.scatter(X,Y)->nuage des points 
plt.show()
plt.close()