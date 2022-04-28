import matplotlib.pyplot as plt
import math


x = [2, 4, 8, 16, 32, 64, 93, 163, 193, 233, 283, 409]

y = []

for i in range(0, len(x)):
    y.append(   ( (x[i]-1)**2 + x[i]**2 ) / 1000   )




# plotting the points
plt.plot(x, y, color='green', linestyle='--', linewidth = 3,
		marker='o', markerfacecolor='black', markersize=12, label='Conventional')







y_KA = []

for i in range(0, len(x)):
    y_KA.append( (7*(x[i]**(math.log(3,2)))-8*x[i]+2)/1000  )




# plotting the points
plt.plot(x, y_KA, color='red', linestyle='--', linewidth = 3,
		marker='o', markerfacecolor='black', markersize=12, label='Karatsuba and Overlap-free')



# naming the x axis
plt.xlabel('Operand Sizes')
# naming the y axis
plt.ylabel('Number of Gates (x1000)')

plt.xticks(x, x)
plt.yticks()










# giving a title to my graph
plt.title('Theoretical calculations')

plt.legend() 

# function to show the plot
plt.show()
