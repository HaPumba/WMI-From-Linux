"""
-WMI from Linux-
* Python2.7

Date	: 10/10/2023
Author	: N.P Matrix

Description: 
This Python script is designed to execute WMI (Windows Management Instrumentation) commands on remote Windows machines from a Linux environment.
This script provides options for specifying the username, password, target IP address or range, and the WMI command to be executed.

Todo list:
V * Add command parameter
V * Validate IP
V * Add IP range to param
V * Optional paramaters
	\ Thred
  	\ Display output

  * Validate operational
V * Execution summary
  * Read command and IP dest from file option
V * Threding
  * Python 3.xx
  * PyInstaller
  * Documentation (including installs)
"""

# Imports
import threading
import argparse as arg
import wmi_client_wrapper as wmi

# Constants
MAX_RANGE = 200

# Globals
exceptionCounter = 0


def ParameterHandler():
	""" This function gathers and returns
	the command line parameters provided to this program """
	
	# Initiate parser
	parser = arg.ArgumentParser()
	
	# Add arguments to parser
	parser.add_argument('-U', '--UserName', type=str, required=True, metavar='[UserName]',
						help="The Username used by the program to execute the WMI command.")
	
	parser.add_argument('-P', '--Password', type=str, required=True, metavar='[Password]',
						help="The Password used by the program to execute the WMI command, '_' (underscore), is used to represent an empty password.")

	parser.add_argument('-D', '--Destenation', type=str, required=True, metavar='[192.168.1.1,199.168.10.10-15]',
						help="The IP address of the destenation machine to execute the WMI command on, this parameter can be a range described using a '-' (minus) or a list separated by ',' (comma).")

	parser.add_argument('-C', '--Command', type=str, required=True, metavar='[WMI command]',
						help="The WMI command to be executed.")
	
	parser.add_argument('-O', action='store_false',
						help="When this flag is included, no output whould be shown.")
	
	parser.add_argument('-T', action='store_false',
						help="When this flag is included, the program won't use threads.")

	arguments = parser.parse_args()
	return(arguments.UserName, arguments.Password, arguments.Destenation, arguments.Command, arguments.O, arguments.T)


def TranslateAddressRange(destenations):
	""" Takes an IP address range and coverts it to a list """
	IPList		= []
	rangeMaker 	= []

	# Test if destenation is a range
	if not "-" in destenations:
		return [destenations]

	# Splits the IP range into octets
	rangeParts = destenations.split('.')

	for part in rangeParts:
		
		# Check if octet contains a rage segment 
		if "-" in part:
			try:
				start, end = part.split("-")

				start, end = int(start),  int(end)

			# Detect the use of invalid range parts
			except Exception:
				print("Bad range parts")
				return False

			# Test if the range was constructed correctlly
			if not start < end:
				print("Bad address range")
				return False

			# Quick test if dest len is going to be bigger than the limit (not conclusive)
			if end - start > MAX_RANGE:
				print("Address range to big")
				return False

		else:
			# Detect the use of invalid range parts
			if not part.isdigit():
				print("Bad address part")
				return False

			start = end = int(part)

		# Test if start and end are in range 
		if start < 0 or end > 255:
			print("Ragne out side of possible octet")
			return False

		# Assigne the the rage start and end location to the range maker
		rangeMaker.append(start)
		rangeMaker.append(end + 1)

	# Checks for an error in the construction of rangeMaker
	if not len(rangeMaker) == 8:
		print("Something went wrong!")
		return False 
	
	# Create IP list based on the ranges stored in the range maker
	for octetA in range(rangeMaker[0], rangeMaker[1]):
		for octetB in range(rangeMaker[2], rangeMaker[3]):
			for octetC in range(rangeMaker[4], rangeMaker[5]):
				for octetD in range(rangeMaker[6], rangeMaker[7]):
					IPList.append("{}.{}.{}.{}".format(octetA, octetB, octetC, octetD))

	if len(IPList) > MAX_RANGE:
		print("Address range to big")
		return False
	
	return IPList


def ValidateIP(ip):
	""" This function "validates" a given IP address """

	# Splits the IP address into octets
	ipParts = ip.split('.')

	# Check if the address has the "right" number of octets
	if len(ipParts) != 4:
		return False

	for octet in ipParts:

		# Check if all octets are valid numbers
		if not octet.isdigit():
			return False

		# Check if all octets are betwenn 0 and 255
		if int(octet) >= 255 or int(octet) < 0:
			return False

	return True

def ExecuteWMICommand(argument):
	global exceptionCounter
	username, password, ip, command, displayOut = argument

	try:
		# # Consructing the WMI connection
		wmiCommand = wmi.WmiClientWrapper(username	= username,
										password 	= password,
										host 		= ip)

		# Executing the command
		output = wmiCommand.query(command)
		
		# Show outpout based on O
		if displayOut:
			print("Executed command on {}\nOutput{}".foramt(ip, output))

	# Error handling
	except Exception as exc:
		
		# Enumerate exception counter
		exceptionCounter += 1

		# Show error based on O
		if displayOut:
			print("Exception occurd while executing the following command on {}: {}\n\n".format(ip, exc))


def main():
	global exceptionCounter
	destList = []

	# Assigne parameters to variables
	username, password, dest, command, output, useThreading = ParameterHandler()
	print("Username: \t{}\nPassword: \t{}\nDestenation: \t{}\nCommand: \t{}".format(username, password, dest, command))

	# Convert _ to empty passwotd
	if password == "_":
		password = ""

	# Create a dest list bsed on ','
	dest = dest.split(",")

	# Convert items in list from rages to addresses
	for section in dest:
		_range = TranslateAddressRange(section)
		
		# Check for errors	
		if _range:

			# Appends the dest IP list with the new IP range
			destList += _range
		else:
			return 0

	# Stors the number of addresses
	destCount = len(destList)

	# Check the number of adresses is smaller than max
	if destCount > MAX_RANGE:
		print("To many addresses")
		return 0

	print("The commands are being executed this may take a while ...")

	# Iterate over all of the given IPs
	for ip in destList:

		# Check the provided IP address
		if not ValidateIP(ip):
			print ("Bad destenation: {}".format(ip))
			continue

		# Test if -T flag was given
		if not useThreading:
			ExecuteWMICommand([username, password, ip, command, output])

		else:
			thread = threading.Thread(target=ExecuteWMICommand, args=([username, password, ip, command, output], ))
			thread.start()

	# Wait for the threads to finish
	thread.join()

	# Test if errors occured
	if exceptionCounter == 0:
		print("Execution is done with no errors!")

	else:
		print("Execution finished with {} errors out of {}".format(exceptionCounter, destCount))

if __name__ == "__main__":
	main()
