from pymouse import PyMouse
import random
import time

m = PyMouse()
a = m.position()
time.sleep(random.randint(5, 10))
m.click(8, 5)
print(a)
