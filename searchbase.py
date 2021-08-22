"""
A Search module for searching the tests to be performed and the cost 
involved from the test and cost databases respectively
"""

__author__ = "Monish Mohanan"
__version__ = "1.0"

# Importing required libraries
try:
	import re
	from tkinter import messagebox
except ImportError as e:
	from tkinter import messagebox
	messagebox.showarning("Import Error", str(e))

class LinearSearch:
	"""
	A class for performing linear search based on the input criteria

	Atributes:
	---------
		change_type : int
			Selected change type in the application

		test_database : xlrd.book.Book object
			Workbook object of the test database file

		cost_database : xlrd.book.Book object
			Workbook object of the cost database file

	Method:
	-------
		extract_test : Extracts the work package ids and test names
		
		extract_cost : Extracts the cost information
	"""

	def __init__(self, change_type, test_database, cost_database):
		"""
		Constructs the required identifiers for initiating the search

		Parameters:
		----------
			change_type : int
				User selected change type

			test_database : xlrd.book.Book object
				Workbook object of the test database file

			cost_database : xlrd.book.Book object
				Workbook object of the cost database file

		Return:
		-------
			None
		"""

		# Assigning the identifiers for linear search
		self.change_type_ = change_type
		self.test_workbook = test_database
		self.cost_workbook = cost_database
		self.change = str(change_type)
		self.test_wpids = list()
		self.test_names = list()
		self.cost_values = list()
		self.test_results = dict()
		self.cost_results = dict()

	def extract_test(self, column):
		"""
		Search the test database based on change type, subassembly and part.
		Return the test names and it's respective work package ids

		Parameters:
		----------
			column : int
				The search column in the test database

		Return:
		-------
			self.test_results : dict
				Contains work package ids & test names as key & value pairs
		"""

		# Assigning the sheet number in the test database
		self.test_sheet = self.test_workbook.sheet_by_index(1)

		# Assigning the rows and columns to search
		self.row = 3
		self.search_column = column - 1

		# Warning message if there are no work package ids
		self.no_wpid = """No tests available for the selection"""

		# Searching for the entire row range in the search column
		for row in range(self.row, self.test_sheet.nrows):
			self.cell_ = str(self.test_sheet.cell_value(row, self.search_column))

			# Searching for the selected type of change in the column
			# Storing the test names and the work package ids separately
			if re.search(rf"{self.change}", self.cell_):
				self.test_wpids.append(
					self.test_sheet.cell_value(row, 0))
				self.test_names.append(
					str(self.test_sheet.cell_value(row, 1)))
			else:
				continue

		# Returning warning message if there aren't any work package ids
		if not bool(self.test_wpids):
			return self.no_wpid

		# Returning warning message if there aren't any test names
		if not bool(self.test_names):
			return "No test data available"

		# Replacing the values as "NULL" if the ids contain inconsistent data
		self.workpackage_ids = list()
		for workpackage_id in self.test_wpids:
			if (isinstance(workpackage_id, float) 
					or len(workpackage_id) == 7):
				self.workpackage_ids.append(str(int(workpackage_id)))
			else:
				self.workpackage_ids.append("NULL")

		# Returning mismatch warning if there is any mismatch in the data
		if not (len(self.workpackage_ids) == len(self.test_names)):
			return "Data mismatch"

		# Assigning work package ids & test names to the test_results
		# as respective key and value pairs
		for key, value in zip(self.workpackage_ids, self.test_names):
			self.test_results[key] = (value.encode("ascii", "ignore")).decode()

		return self.test_results

	def extract_cost(self, wp_ids):

		"""
		Search for costs in the cost database based on work package ids

		Parameters:
		----------
			wp_ids : dict_keys
				Collection of work package ids of the tests

		Return:
		------
			self.cost_results : dict
				Contains workpackage ids and their respective costs as
				key & value pairs 
				ids - string and costs - float
		"""

		self.wp_ids_ = wp_ids

		self.cost_data = {}

		for sheet in self.cost_workbook.sheets():

			for row in range(1, sheet.nrows):
				package = str(sheet.cell_value(row, 2))
				cost = sheet.cell_value(row, 16)

				if package != None and package != "":
					try:
						self.cost_data[package] = round(float(cost), 1)
					except:
						continue

		for test in self.wp_ids_:
			try:
				self.cost_results[test] = self.cost_data[test]
			except:
				continue

		if len(self.wp_ids_) != len(self.cost_results):
			return "Missing workpackage/cost info in cost database"
		else:
			return self.cost_results
