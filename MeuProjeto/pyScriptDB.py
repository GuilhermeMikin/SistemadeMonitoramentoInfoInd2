import time
import datetime 

date = str(datetime.datetime.fromtimestamp(int(time.time())).strftime("%Y-%m-%d %H:%M:%S"))

file = open("datevalue.txt","w") 
file.write(date) 
file.close() 

