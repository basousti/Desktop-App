import matplotlib.pyplot as plt
x = [0, 1, 0] 
y = [0, 1, 2] 
x1 = [0, 2, 0] 
y1 = [2, 1, 0] 
x2 = [0, 1, 2] 
y2 = [0, 1, 2]
plt.title("titre")
plt.plot(x,y,color="black",label="F1")
plt.plot( x1, y1,lw=5,label="F2")
plt.plot(x2, y2,label="F3") #les axes
plt.xlabel( "abscisse") 
plt.ylabel("ordonnees")
plt.legend()
plt.show ()
plt.close()