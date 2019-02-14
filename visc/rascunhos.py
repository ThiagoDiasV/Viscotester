import pandas as pd 
from sklearn import linear_model

dataframe = pd.DataFrame()
dataframe['x'] = [1, 2, 3, 4, 5]
dataframe['y'] = [100, 250, 300, 400, 500]
x_values = dataframe[['x']]
y_values = dataframe[['y']]

model = linear_model.LinearRegression()
model.fit(x_values, y_values)

print('y = ax + b')
print(model.score(x_values, y_values))
print('a = %.2f = a inclinação da linha de tendência.' % model.coef_)
print('b = %.2f = o ponto onde a linha de tendência atinge o eixo y.' % model.intercept_)