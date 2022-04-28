import matplotlib.pyplot as plt
import math


x = [2, 4, 8, 16, 32, 64, 93, 163, 193, 233, 283, 409]

y = []

for i in range(0, len(x)):
    y.append( 1+math.log(x[i],2)  )




# plotting the points
plt.plot(x, y, color='green', linestyle='--', linewidth = 3,
		marker='o', markerfacecolor='black', markersize=12, label='Conventional')








y_KA = []

for i in range(0, len(x)):
    y_KA.append( 3*math.log(x[i],2)  )




# plotting the points
plt.plot(x, y_KA, color='red', linestyle='--', linewidth = 3,
		marker='o', markerfacecolor='black', markersize=12, label='Karatsuba')







y_OKA = []

for i in range(0, len(x)):
    y_OKA.append( 2*math.log(x[i],2)  )




# plotting the points
plt.plot(x, y_OKA, color='blue', linestyle='--', linewidth = 3,
		marker='o', markerfacecolor='black', markersize=12, label='Overlap-free')



# naming the x axis
plt.xlabel('Operand Sizes')
# naming the y axis
plt.ylabel('Delay (Tg)')

plt.xticks(x, x)
plt.yticks()










# giving a title to my graph
plt.title('Theoretical calculations')

plt.legend() 

# function to show the plot
plt.show()
