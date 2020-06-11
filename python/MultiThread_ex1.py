#!/usr/bin/python

import thread
import time

# Define a function for the thread
def print_time( threadName, delay, number):
   count = 0
   while count < number:
      time.sleep(delay)
      count += 1
      print "%s.%s: %s" % ( threadName,count, time.ctime(time.time()) )

# Create two threads as follows
try:
   thread.start_new_thread( print_time, ("Thread-1", 2, 5,) )
   thread.start_new_thread( print_time, ("Thread-2", 4, 5,) )
except:
   print "Error: unable to start thread"

while 1:
   pass
