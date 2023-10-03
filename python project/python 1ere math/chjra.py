import matplotlib.pyplot as plt
x = [0.25, 0.25, 1.25, 0.5, 1, 0.25, 0.6, 0, -0.6, -0.25, -1, -0.5,-1.25, -0.25, -0.25, 0.25]
y = [0, 0.5, 0.5, 1, 1, 1.5, 1.5, 2, 1.5 , 1.5, 1, 1, 0.5, 0.5, 0,0]
plt.plot(x, y, '-.', color = "green", lw = 2)
plt.title("Dessin") 
plt.axis('equal') 
plt.xlabel("arbre avec python")
plt.ylabel("Vive le vent") 
plt.show() 
plt.close() 
