import matplotlib.pyplot as plt
import numpy as np
a = 1 
b = 20 # On choisit de dessiner sur [1, 20]
x = [] 
y = [] 
pas = (b - a) / 200 # On choisit de diviser l'intervalle en 200
abscisse = a # L'abscisse de départ est a cette fois
for k in range(0, 201): # On va donc de 0 à 200
 x.append(abscisse) 
 y.append(np.log(abscisse)) 
 abscisse += pas 
plt.plot(x, y) 
plt.title("logarithme") 
plt.show() 
plt.close()