from csv import DictReader
import os

#os.name function gives the name of the operating system dependent module imported. The following names have currently been registered: 
#'posix', 'nt', 'os2', 'ce', 'java' and 'riscos'

print(os.name) 

fd = "GFG.txt"

#file = open(fd, 'w')
#file.write("Hello")
#file.close()
#file = open(fd, 'r')
#text = file.read()
#print(text)

file = os.popen(fd, 'w')
file.write("Hello")


#Changing the current working directory
def current_path():
    #Get the current working directory (CWD)
    cwd = os.getcwd()
    print("Current working directory before", cwd)

#Changing the CWD
current_path()
os.chdir('../')
current_path()

#Creating a directory with using os.mkdir() used to create a directory named path with the specified numeric mode. 
#This method raises FileExistsError if the directory to be created already exists. 

parent_dir = "C:/Users/long.pham/Documents/MDDPython/examples"

directory = "os module 1"

try:
    path = os.path.join(parent_dir, directory)

    os.mkdir(path)

    print("Directory '% s' created" % directory)
except FileExistsError as ex:
    print("Directory '% s' to be created already exists." % directory)

#Creating a directory with using os.makedirs() used to create a directory recursively. That means while making leaf directory
#if any intermediate-level directory is missing, os.makedirs() method will them all.

parent_dir = "C:/Users/long.pham/Documents/MDDPython/examples/a/b"

directory = "os module 2"

try:
    path = os.path.join(parent_dir, directory)

    os.makedirs(path)

    print("Directory '% s' created" % directory)
except FileExistsError as ex:
    print("Directory '% s' to be created already exists." % directory)

#Get a list of all files and directories in a root directory
path = "/"

dir_list = os.listdir(path)

print(dir_list)

#os.remove() method in python used to remove or delete a file path. This method can not remove a directory. If a specified path 
#is a directory then OSError will be raised by the method.

file = "file1.txt"

location = "C:/Users/long.pham/Documents/MDDPython/examples"

path = os.path.join(location, file)

try:
    os.remove(path)
except FileNotFoundError as ex:
    print("The system cannot find the file specified: '% s'" % path)
except OSError as ex:
    print("Access is denied: '% s'" % path)

#os.rmdir() used to remove or delete an empty directory. OSError will be raised if the specified path is not an empty directory.

directory = "os module 1"

parent_dir = "C:/Users/long.pham/Documents/MDDPython/examples"

path = os.path.join(parent_dir, directory)

try:
    os.rmdir(path)
except OSError as ex:
    print("The system cannot delete an empty directory: '% s'" % path)

