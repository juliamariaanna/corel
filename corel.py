import xlrd, numpy, re, sklearn.metrics, sklearn.linear_model, prettytable, os
# Open file
file=xlrd.open_workbook("./data.xls")
sheet_index=None
row_index=None
flag=False
for i in range(0, file.nsheets, 1):
    sheet=file.sheet_by_index(i)
    rows=sheet.nrows
    cols=sheet.ncols
    for j in range(0, rows, 1):
        for k in range(0, cols, 1):
            if str(sheet.cell_value(j,k)) == 'Ukraine':
                sheet_index=i
                row_index=j
                flag=True
                break
            if flag:
                break
        if flag:
            break
    if flag:
        break
# Collect data
X=[]
X_LAG=[]
sheet=file.sheet_by_index(sheet_index)
cols=sheet.ncols
num_pattern=re.compile(r"^[\d\.]+$")
for i in range(0, cols, 1):
    if re.match(num_pattern, str(sheet.cell_value(row_index, i))) != None:
        l=len(X)
        X.append(float(sheet.cell_value(row_index, i)))
        if l > 0:
            X_LAG.append(float(sheet.cell_value(row_index, i)))
X.pop()
# Calculate XY
XY=[]
X2=[]
X_LAG2=[]
for i in range(0, len(X), 1):
    XY.append(X[i]*X_LAG[i])
    X2.append(X[i] * X[i])
    X_LAG2.append(X_LAG[i] * X_LAG[i])
# Calculate Pearson correlation coefficient
# print("Correl by numpy: " + str(numpy.corrcoef(X,X_LAG)[0][1]))
numerator=( (sum(XY)/len(XY))-(sum(X)/len(X))*(sum(X_LAG)/len(X_LAG)) )
Xa=(sum(X2)/len(X2) - (sum(X)/len(X))*(sum(X)/len(X)))
X_LAGa=(sum(X_LAG2)/len(X_LAG2) - (sum(X_LAG)/len(X_LAG))*(sum(X_LAG)/len(X_LAG)))
# print("My own calculation: " + str(numerator/pow(Xa*X_LAGa,0.5)))
# Coefitient of determination
x = numpy.array(X).reshape((-1, 1))
y = numpy.array(X_LAG)

# print(f"R^2 = {sklearn.metrics.r2_score(X, X_LAG)} by sklearn")
# Model (calculated by myself)
k = (sum(X) * sum(X_LAG) - len(XY) * sum(XY))/(sum(X) ** 2 - len(X2) * sum(X2))
b = (sum(X_LAG) - k * sum(X))/len(X)
# print(f'Own calculation\nf(x) = {k}*x + {b}')
# Model (calculated by sklearn)
model = sklearn.linear_model.LinearRegression().fit(x, y)
# print(f'sklearn calculation\nf(x) = {model.coef_[0]}*x + {model.intercept_}')
# print(f'R^2 = {model.score(x, y)} calculated by sklearn after build model')
# Calculate predicted data
RSS = 0
TSS = 0
X_LAGi = []
for i in range(0, len(X)):
    X_LAGi.append(k * X[i] + b)
    RSS += ((k * X[i] + b) - sum(X_LAG)/len(X_LAG))**2
    TSS += (X_LAG[i] - sum(X_LAG)/len(X_LAG))**2

# R^2 calculated by myself
# print(f'R^2 = {RSS/TSS} calculated by myself')
# Build table
pt = prettytable.PrettyTable()
pt.field_names = ['X', 'X_LAG', 'X_LAGi']
for i in range(0, len(X)):
    pt.add_row([X[i], X_LAG[i], X_LAGi[i]])
# Save result
output = open('corel_output.txt', 'a') if not os.path.exists('./corel_output.txt')\
      else open('corel_output.txt', 'w')
output.write(str(pt))
output.write(f'\n\n\nf(x) = {k}*x + {b}\n')
output.write(f'R^2 = {RSS/TSS}')
