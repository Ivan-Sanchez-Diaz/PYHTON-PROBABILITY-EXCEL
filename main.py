import os
import openpyxl
os.system("cls")

example = [2.97,4.00,5.20,5.56,5.94,5.98,6.35,6.62,6.72,6.78,6.80,6.85,6.94,7.15,7.16,7.23,7.29,7.62,7.62,7.73,7.87,7.93,8.00,8.26,8.29,8.37,8.47,8.54,8.54,8.58,8.61,8.67,8.69,8.81,9.07,9.27,9.37,9.43,9.52,9.58,9.60,9.76,9.82,9.83,9.83,9.84,9.96,10.04,10.21,10.28,10.28,10.30,10.35,10.36,10.40,10.49,10.49,10.50,10.64,10.95,11.09,11.12,11.21,11.29,11.43,11.62,11.70,12.16,12.19,12.28,12.31,12.62,12.69,12.71,12.91,12.92,13.11,13.38,13.42,13.43,13.47,13.60,13.96,14.24,14.35,15.12,15.24,16.06,16.90,18.26]
example_2 = [11.5,12.1,9.9,9.3,7.8,6.2,6.6,7.0,13.4,17.1,9.3,5.6,5.7,5.4,5.2,5.1,4.9,10.7,15.2,8.5,4.2,4.0,3.9,3.8,3.6,3.4,20.6,25.5,13.8,12.6,13.1,8.9,8.2,10.7,14.2,7.6,5.2,5.5,5.1,5.0,5.2,4.8,4.1,3.8,3.7,3.6,3.6,3.6]
def data(array):
  total = len(array)
  array.sort()
  interval = array[0] - array[total-1]
  #print(array)

  #print("total:"+str(total))
  #print("interval:" +str(abs(interval)))

  classNumber = 6
  ValueNumber = 4
  off = 0
  initial = 0
  y = 0

  classInterval = []
  middlePoint = []
  classFrequency = []
  relativeFrequency = []
  display = []

  for i in range(classNumber):
    if off>0:
      initial = array[0]/1000

    parameter = (array[0]+off+initial),(array[0]+ValueNumber+off)
    classInterval.append(parameter)
    off = off + ValueNumber

  for i in classInterval:
    middlePoint.append((i[0]+i[1])/2)
    count = 0
    for j in array:
      #print("j:"+str(j))
      #print("("+str(i[0])+","+str(i[1])+")")
      if( j >= i[0] and j<=i[1]):
        #print(str(j)+","+str(y))
        #print(str("(")+str(j)+str(")")+"*")
        count = count + 1
    #print(count)
    y = y + 1
    classFrequency.append(count)
  
  for i in classFrequency:
    relativeFrequency.append(i/total)

  #print(classInterval)
  #print(classFrequency)
  #print(relativeFrequency)

  display.append(classInterval)
  display.append(classFrequency)
  display.append(relativeFrequency)
  
wb = openpyxl.Workbook()



if __name__ == "__main__":
  data(example)