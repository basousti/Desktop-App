import matplotlib.pyplot as plt
a=1/1000#pour eviter le 0
b=1
X=[]
Y=[]
X1=[]
Y1=[]
pas=(b-a)/200
abscisse =a
for k in range (0,201):
    X.append(abscisse)
    Y.append(1/abscisse)
    X.append(-abscisse)
    Y.append(-1/abscisse)
    abscisse+=pas
plt.title("etude de fonction ln")
plt.plot(X,Y,color="pink") #plt.scatter(X,Y)->nuage des points 
plt.show()
plt.close()