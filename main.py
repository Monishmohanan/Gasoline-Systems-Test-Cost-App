"""
A master file for handling the user interface and the process flow
of the application.

	Dependencies:
	------------
		info.db - Contains all the required information for the app
		report.db - Contains usage information of the application

	Windows:
	-------
		MainWindow - For obtaining input combination from the user
		ConfirmationWindow - For confirming the procured results
		Settings - For modifying paths, users & report generation
		
"""

__author__ = "Monish Mohanan"
__version__ = "1.0"

try:
	import os
	import sys
	import sqlite3
	import logging
	import traceback
	import time
	import xlrd
	from concurrent.futures import ThreadPoolExecutor
	from PIL import ImageTk, Image
	import tkinter as tk
	import tkinter.scrolledtext as tkst
	from tkinter import messagebox
	from tkinter import ttk
	from tkinter.filedialog import askopenfile
	from babel.numbers import format_currency
	from win32com import client
	from report import TransmissionReport
	from searchbase import LinearSearch
	from template import TransmissionTemplate
except ImportError as e:
	from tkinter import messagebox
	messagebox.showwarning("Import Error", str(e))
	sys.exit(0)

try:
	logging.basicConfig(
		filename='TestCostApp.log', 
		format='%(asctime)s - %(message)s', 
		datefmt='%d-%b-%y %H:%M:%S')
except Exception as e:
	raise e

try:
	info = sqlite3.connect('database/info.db')
	cur = info.cursor()
	cur.execute('''SELECT Number, Changes FROM ChangeTypes''')
	change_types = dict((col[0], col[1]) for col in cur.fetchall())
	cur.execute('''SELECT Name, Position FROM subAssembly''')
	subassemblies = dict((col[0], col[1]) for col in cur.fetchall())
	cur.execute('''SELECT PartName, Position FROM Subassembly_1''')
	subassembly_1 = dict((col[0], col[1]) for col in cur.fetchall())
	cur.execute('''SELECT PartName, Position FROM Subassembly_2''')
	subassembly_2 = dict((col[0], col[1]) for col in cur.fetchall())
	cur.execute('''SELECT PartName, Position FROM Subassembly_3''')
	subassembly_3 = dict((col[0], col[1]) for col in cur.fetchall())
	cur.execute('''SELECT PartName, Position FROM Subassembly_4''')
	subassembly_4 = dict((col[0], col[1]) for col in cur.fetchall())
	cur.execute('''SELECT PartName, Position FROM Subassembly_5''')
	subassembly_5 = dict((col[0], col[1]) for col in cur.fetchall())
	cur.execute('''SELECT PartName, Position FROM Subassembly_6''')
	subassembly_6 = dict((col[0], col[1]) for col in cur.fetchall())
	cur.execute('''SELECT Name FROM Requesters''')
	requesters = [col[0] for col in cur.fetchall()]
	cur.execute('''SELECT Name FROM Creators''')
	creators = [col[0] for col in cur.fetchall()]
	cur.execute('''SELECT Name, Path FROM Databases''')
	databases = dict((col[0], col[1]) for col in cur.fetchall())
	cur.execute('''SELECT Name, Path FROM Storage''')
	records = dict((col[0], col[1]) for col in cur.fetchall())
except Exception as e:
	messagebox.showwarning("SQL Query Error", str(e))
	logging.error(traceback.format_exc())
	sys.exit(0)
else:
	partlist = (
		subassembly_1, subassembly_2,
		subassembly_3, subassembly_4,
		subassembly_5, subassembly_6,
		)
	subassembly_keys = list(subassemblies.keys())
	keys = (0, 1, 2, 3, 4, 5)
	required_keys = tuple((subassembly_keys[val]) for val in keys)
	subassembly_and_parts = dict(
		(key, value) for key, value in zip(required_keys, partlist))
finally:
	info.close()

MAIN_WINDOW_TITLE = "TEST AND COST TEMPLATE"
MAIN_WINDOW_RESOLUTION = "850x650"
CONFIRMATION_WINDOW_TITLE = "CONFIRMATION SCREEN"
CONFIRMATION_WINDOW_RESOLUTION = "800x650"
SETTINGS_WINDOW_TITLE = "SETTINGS"
SETTINGS_WINDOW_RESOLUTION = "550x500"
DATA_ADDITION_WINDOW_TITLE = "ADD FIELDS"
DATA_ADDITION_WINDOW_RESOLUTION = "300x130"
DATA_DELETION_WINDOW_TITLE = "REMOVE FIELDS"
DATA_DELETION_WINDOW_RESOLUTION = "450x250"
TITLE = "TEST AND COST TEMPLATE GENERATOR"
IMAGE_TITLE = "Transmission System"
PRODUCT_IMAGE = "images/Transmission.jpg"
PRODUCT_LOGO = "images/mechanism.png"
DOCS_IMAGE = "images/docs.png"
SETTINGS_IMAGE = "images/settings.png"
TEMPLATE_IMAGE = "images/folder.png"
REPORT_IMAGE = "images/report.png"
TEMPLATE_PATH = "templates/"
REPORT_PATH = "report/"



class MainWindow:
	"""
	A class to represent the primary window of the application.

	...

	Attributes:
	----------
		master : tktiner.Tk class
			Base class for construction of the main window

	Method:
	-------
		confirmation_window : Instantiates confirmation window

		workflow : Triggers workflow based on input validation

		validate_inputs : Validates the input recieved from the user

		on_subassembly_change : Subassembly selection in the application

		on_part_change : Part name selection in the application

		documentation : Opens the documentation

		settings : Instantiates the settings window of the application
	"""

	def __init__(self, master):
		"""
		Constructs the user interface for the main window

		Parameters:
		----------
			master : tkinter.Tk class
				Base class for construction of the main window
		"""

		# Basic configuration of the window
		self.master = master
		self.master.title(MAIN_WINDOW_TITLE)
		self.master.geometry(MAIN_WINDOW_RESOLUTION)
		self.master.resizable(0, 0)
		self.master.configure(background = 'white')

		# Defining input variable wrappers
		self.change_type = tk.IntVar()
		self.subassembly = tk.StringVar()
		self.part = tk.StringVar()
		self.requester = tk.StringVar()
		self.creator = tk.StringVar()

		# Adding Bosch logo
		self.product_logo = ImageTk.PhotoImage(Image.open(PRODUCT_LOGO))
		tk.Label(image = self.product_logo, bg = 'white').place(x = 40, y = 0)

		# Adding product image & text
		self.product_image = ImageTk.PhotoImage(Image.open(PRODUCT_IMAGE))
		tk.Label(image = self.product_image, bg = 'white').place(x = 580, y = 180)
		tk.Label(self.master, text = IMAGE_TITLE, font = ('Arial bold', 15), 
			fg = 'black', bg = 'white').place(x = 600, y = 450)

		# Window title
		tk.Label(self.master, text = TITLE, font = ('Arial bold', 20), 
			fg = 'dark blue', bg = 'white').place(x = 100, y = 12)

		# Documentation & settings buttons
		self.docs_image = ImageTk.PhotoImage(Image.open(DOCS_IMAGE))
		self.settings_image = ImageTk.PhotoImage(Image.open(SETTINGS_IMAGE))
		self.template_image = ImageTk.PhotoImage(Image.open(TEMPLATE_IMAGE))
		self.report_image = ImageTk.PhotoImage(Image.open(REPORT_IMAGE))
		tk.Button(self.master, image = self.settings_image, bg = 'white',
			command = self.settings).place(x = 800, y = 15)
		tk.Button(self.master, image = self.docs_image, bg = 'white', 
			command = self.documentation).place(x = 800, y = 60)

		self.file_path = os.path.join(os.getcwd(), TEMPLATE_PATH)
		tk.Button(self.master, image = self.template_image, bg = 'white',
			command = lambda:os.startfile(self.file_path)).place(x = 800, y = 105)

		self.report_path = os.path.join(os.getcwd(), REPORT_PATH)
		tk.Button(self.master, image = self.report_image, bg = 'white',
			command = lambda:os.startfile(self.report_path)).place(x = 800, y = 150)

		# ------------------------CHANGE TYPE LAYOUT----------------------------

		self.pos = 0
		tk.Frame(
			self.master,
			width = 515,
			height = 162,
			background = 'white', 
			highlightthickness = 4).place(x = 40, y = 65)
		tk.Label(
			self.master, text = "TYPE OF CHANGE", 
			bg = 'white', fg = 'black', 
			font = ('Arial bold', 15)).place(x = 45, y = 55)
		for number, changes in change_types.items():
			tk.Radiobutton(
				self.master, text = ' '.join((str(number), changes)),
				bg = 'white', fg = 'black', 
				font = ('Times New Roman', 13),
				variable = self.change_type, 
				value = number).place(x = 50, y = 95 + 30 * self.pos)
			self.pos += 1

		# -----------------------SUB ASSEMBLY LAYOUT---------------------------

		tk.Frame(
			self.master,
			width = 515,
			height = 70,
			background = 'white',
			highlightthickness = 4).place(x = 40, y = 245)
		tk.Label(
			self.master, text = "SUB ASSEMBLY",
			bg = 'white', fg = 'black',
			font = ('Arial bold', 15)).place(x = 45, y = 235)
		tk.Label(
			self.master, text = "Select the Subassembly: ",
			bg = 'white', fg = 'black',
			font = ('Times New Roman', 13)).place(x = 50, y = 275)
		self.subassembly_dropmenu = tk.OptionMenu(
			self.master, self.subassembly,
			*list(subassemblies.keys()), command = self.on_subassembly_change)
		self.subassembly_dropmenu.config(
			bg = 'white', fg = 'dark blue', 
			width = 35, relief = tk.GROOVE)
		self.subassembly_dropmenu.place(x = 230, y = 273)

		# --------------------------PART LAYOUT--------------------------------

		tk.Frame(
			self.master,
			width = 515,
			height = 70,
			background = 'white',
			highlightthickness = 4).place(x = 40, y = 340)
		tk.Label(
			self.master, text = "PART NAME",
			bg = 'white', fg = 'black',
			font = ('Arial bold', 15)).place(x = 45, y = 330)
		tk.Label(
			self.master, text = "Select the part name: ",
			bg = 'white', fg = 'black',
			font = ('Times New Roman', 13)).place(x = 50, y = 370)
		self.part_dropmenu = tk.OptionMenu(
			self.master, self.part,
			'', command = self.on_part_change)
		self.part_dropmenu.config(
			bg = 'white', fg = 'dark blue',
			width = 30, relief = tk.GROOVE)
		self.part_dropmenu.place(x = 230, y = 367)

		# -----------------------ADDITIONAL FIELDS----------------------------

		tk.Frame(
			self.master,
			width = 515,
			height = 150,
			background = 'white',
			highlightthickness = 4).place(x = 40, y = 435)
		tk.Label(
			self.master, text = "ADDITIONAL INFO",
			bg = 'white', fg = 'black',
			font = ('Arial bold', 15)).place(x = 45, y = 425)
		tk.Label(
			self.master, text = "Requester: ",
			bg = 'white', fg = 'black',
			font = ('Times New Roman', 13)).place(x = 50, y = 455)
		ttk.Combobox(
			self.master, values = requesters,
			width = 45, foreground = 'dark blue',
			state = "readonly",
			textvariable = self.requester).place(x = 230, y = 455)
		tk.Label(
			self.master, text = "Creator: ",
			bg = 'white', fg = 'black',
			font = ('Times New Roman', 13)).place(x = 50, y = 485)
		ttk.Combobox(
			self.master, values = creators,
			width = 45, foreground = 'dark blue',
			state = "readonly",
			textvariable = self.creator).place(x = 230, y = 485)
		tk.Label(
			self.master, text = "Comment: ",
			bg = 'white', fg = 'black',
			font = ('Times New Roman', 13)).place(x = 50, y = 515)
		self.comment = tk.Text(
			self.master, 
			width = 35, 
			bd = 2.5, 
			height = 2.5)
		self.comment.place(x = 230, y = 518)

		tk.Button(
			self.master,
			text = "Generate",
			font = ('Arial bold', 15),
			bg = 'light gray',
			fg = 'black',
			relief = tk.RAISED,
			activebackground = '#24025F',
			bd = 6, command = self.validate_inputs).place(x = 245, y = 590)

	def confirmation_window(self, test, cost, **kwargs):
		"""
		Instantiate the confirmation window from the 

		tkinter Toplevel class

		Parameters:
		----------
			change : int
				Selected change type in the application

			subassy : str
				Selected subassembly in the application

			part : str
				Selected part in the application

			test : dict
				Work package IDs and test names

			cost : dict
				Work package IDs and costs

		Return:
		------
			None
		"""

		self.confirmation = tk.Toplevel(self.master)
		self.app = ConfirmationWindow(self.confirmation, test,
										cost, **kwargs)

	def workflow(self):
		"""
		Search for tests and costs in the test and cost database

		Invoke the confirmation screen after validation

		Parameters:
		----------
			None

		Return:
		------
			None

		"""

		# Assigning identifiers for searching the databases
		change_type = self.change_type.get()
		subassembly = self.subassembly.get()
		part = self.part.get()
		requester = self.requester.get()
		creator = self.creator.get()
		comment = self.comment.get("1.0", "end-1c")
		if not comment:
			comment = "None"
		self.test_valid = False
		self.cost_valid = False
		self.no_test_message = "No tests for the selected combination"
		self.critical = "Something is wrong. Please contact the developer"
		self.missing_wp = "Missing workpackage IDs in test database"

		# Set the validation criteria for test and cost database loading
		self.load_validation = bool(
			test_database.result()
			and cost_database.result())

		# Verify if the databases are loaded in the application
		while not self.load_validation: print("loading database...")

		# Assign the test and cost database workbook objects to identifiers
		self.test_database = test_database.result()
		self.cost_database = cost_database.result()

		# Set the search column for the test database
		if change_type in range(1, 5):
			if bool(part):
				self.search_column = subassembly_and_parts[subassembly][part]
			else:
				part = "NA"
				self.search_column = subassemblies[subassembly]

		# Instantiate LinearSearch and extract the test results
		search = LinearSearch(
			change_type, 
			self.test_database, 
			self.cost_database)
		test_results = search.extract_test(self.search_column)

		# Validate if the results contain the correct test data
		# Set the validation flag based on the condition
		if isinstance(test_results, dict):
			if test_results:
				if "NULL" not in test_results.keys():
					self.test_valid = True
				else:
					messagebox.showwarning(
						"Insufficient Info", 
						self.missing_wp)
			else:
				messagebox.showwarning("No Data", self.no_test_message)
		elif isinstance(test_results, str):
			messagebox.showwarning("Warning", test_results)
		else:
			messagebox.showwarning("Error", self.critical)
			logging.error(self.critical)
			sys.exit(0)

		# Extract the costs if the test data is valid
		if self.test_valid:
			cost_results = search.extract_cost(test_results.keys())

			# Validate if the results contain the correct cost data
			# Set the validation flag based on the condition
			if isinstance(cost_results, dict):
				if cost_results:
					self.cost_valid = True
				else:
					messagebox.showwarning(
						"No Data",
						"No data recieved from cost database")
			elif isinstance(cost_results, str):
				messagebox.showwarning("Warning", cost_results)
			else:
				messagebox.showwarning("Error", self.critical)
				logging.error(self.critical)
				sys.exit(0)


		# Invoke the confirmation screen if the tests & costs are validated
		if (self.test_valid and self.cost_valid):
			self.inputs = {
				"Change Type" : change_types[change_type],
				"Subassembly" : subassembly,
				"Part Name" : part,
				"Requester" : requester,
				"Creator" : creator,
				"Comment" : comment
				}
			MainWindow.confirmation_window(self, test_results, 
											cost_results, **self.inputs)

	def validate_inputs(self):
		"""
		Validate the inputs, set the validate flag & trigger the workflow

		If the minimum input requirements are not met, warn the user

		Parameters:
		----------
			None

		Return:
		------
			None
		"""

		# Initialise the validation flag and set the change variable
		validate = False
		characters_exceeded = False
		change = self.change_type.get()

		# Set the validation criteria
		self.subassembly_vaidation = bool(self.subassembly.get())
		self.users_validation = bool(
			self.requester.get() and self.creator.get())
		self.comment_validation = self.comment.get("1.0", "end-1c")
		self.limit_warning = "Maximum comment length is 85 characters"

		# Warning message
		self.message = "Please provide all the inputs"

		# Evaluation of change type, subassembly and part input fields
		if change in range(1, 5):
			if self.subassembly_vaidation:
				validate = True
			else:
				validate = False
		else:
			validate = False

		# Evaluation of requester and creator fields
		if not self.users_validation:
			validate = False

		if len(self.comment_validation) > 85:
			characters_exceeded = True

		# Trigger workflow or raise warning based on evaluation critera
		if validate and not characters_exceeded:
			MainWindow.workflow(self)
		elif characters_exceeded:
			messagebox.showwarning("Limit Exceeded", self.limit_warning)
		else:
			messagebox.showwarning("Insufficient Inputs", self.message)





	def on_subassembly_change(self, selection):
		"""
		Get the subassembly input and search for the parts

		Populate the PART NAME field based on search results

		Parameters:
		----------
			selection : str
				subassembly selected by the user

		Returns:
		-------
			None
		"""

		self.menu = self.part_dropmenu['menu']
		self.menu.delete(0, 'end')

		if selection == subassembly_keys[0]:
			self.part_1 = subassembly_and_parts[subassembly_keys[0]]
			self.selected_parts = list(self.part_1.keys())
		elif selection == subassembly_keys[1]:
			self.part_2 = subassembly_and_parts[subassembly_keys[1]]
			self.selected_parts = list(self.part_2.keys())
		elif selection == subassembly_keys[2]:
			self.part_3 = subassembly_and_parts[subassembly_keys[2]]
			self.selected_parts = list(self.part_3.keys())
		elif selection == subassembly_keys[3]:
			self.part_4 = subassembly_and_parts[subassembly_keys[3]]
			self.selected_parts = list(self.part_4.keys())
		elif selection == subassembly_keys[4]:
			self.part_5 = subassembly_and_parts[subassembly_keys[4]]
			self.selected_parts = list(self.part_5.keys())
		elif selection == subassembly_keys[5]:
			self.part_6 = subassembly_and_parts[subassembly_keys[5]]
			self.selected_parts = list(self.part_6.keys())
		else:
			self.selected_parts = ['']

		self.part.set('')

		for item in self.selected_parts:
			self.menu.add_command(
				label = item, 
				command = lambda x = item: self.on_part_change(x))

	def on_part_change(self, selected):
		"""
		Set the part name as per the user selection in the window

		Parameters:
		----------
			selected : str
				part name selected by the user

		Return:
		------
			None
		"""

		self.part.set(selected)


	def documentation(self):
		try:
			os.startfile('Test_Cost_App-docs.pdf')
		except Exception as e:
			messagebox.showwarning(
			"Not available", 
			"Oops! Documentation cannot be accessed")
			logging.error(traceback.format_exc())

	def settings(self):
		self.settings = tk.Toplevel(self.master)
		self.settingsapp = Settings(self.settings)

	@staticmethod
	def load_databases(value):
		"""
		Load the test and the cost database in memory by creating objects

		Parameters:
		----------
			value : str
				Location of the database file

		Return:
		------
			wb : xlrd.book.Book object
				Workbook object of the database file
		"""

		# Try loading the database file 
		try:
			wb = xlrd.open_workbook(value)
		except Exception as e:
			messagebox.showwarning("Database Error", str(e))
			logging.error(traceback.format_exc())
			sys.exit(0)

		# Return the workbook object if the load is successful
		return wb


class ConfirmationWindow:
	"""
	A class to represent the confirmation window of the application

	Attributes:
	----------
		master : tkinter.Tk class
			Base class for the construction of the confirmation window

	Method:
	------
		generate_pdf : Generates the test and cost template
	"""

	def __init__(self, master, test, cost, **kwargs):
		"""
		Constructs the confirmation window of the application

		Parameters:
		----------
			master : tkinter.Tk class
				Base class for the construcation of the window

			test : dict
				Collection of work pacakge ids and test pairs

			cost : dict
				Collection of work package ids and cost pairs

			**kwargs : dict
				Contains change type, subassembly, part name,
				requester, creator and comment values
		"""

		# Basic configuration of the window
		self.master = master
		self.master.title(CONFIRMATION_WINDOW_TITLE)
		self.master.geometry(CONFIRMATION_WINDOW_RESOLUTION)
		self.master.resizable(0, 0)
		self.master.configure(background = 'white')

		# Assigning the required identifiers
		self.test_ = test
		self.cost_ = cost
		self.input_values = kwargs

		# ------------------------------TITLE--------------------------------

		tk.Frame(
			self.master,
			width = 780,
			height = 46,
			background = '#24025F',
			highlightthickness = 4).place(x = 10, y = 0)
		tk.Label(
			self.master, text = "TRANSMISSION SYSTEMS",
			fg = 'white', bg = '#24025F',
			font = ('arial bold', 20)).place(x = 220, y = 4)

		# ------------------------INPUT SELECTION----------------------------

		tk.Frame(
			self.master,
			width = 780,
			height = 46,
			background = '#24025F',
			highlightthickness = 4).place(x = 10, y = 46)
		tk.Label(
			self.master, text = "INPUT SELECTION",
			fg = 'white', bg = '#24025F',
			font = ('Times New Roman', 15)).place(x = 20, y = 55)
		tk.Frame(
			self.master,
			width = 780,
			height = 165,
			background = 'white',
			highlightthickness = 4).place(x = 10, y = 92)

		# Displaying the input selections made by the user
		spacing  = 0
		for key, value in self.input_values.items():
			tk.Label(
				self.master, text = key,
				bg = 'white', fg = 'black',
				font = ('arial bold', 13)).place(x = 20, y = 100 + 25 * spacing)
			tk.Label(
				self.master, text = " ".join((":",value)),
				bg = 'white', fg = 'dark green',
				font = ('helvetica', 11)).place(x = 170, y = 100 + 25 * spacing)
			spacing += 1

		# ----------------------TEST AND COST DATA---------------------------

		tk.Frame(
			self.master,
			width = 780,
			height = 46,
			background = "#24025F",
			highlightthickness = 4).place(x = 10, y = 257)
		tk.Label(
			self.master, text = "TEST AND COST DATA",
			fg = 'white', bg = '#24025F',
			font = ('Times New Roman', 15)).place(x = 20, y = 266)

		# Pairing the test names with the appropriate costs
		self.test_cost = tuple(
			(x, y) for x, y in zip(self.test_.values(),self.cost_.values()))

		# Calculating the total cost and rounding it upto 2 decimal values
		self.total_cost = [float(num) for num in cost.values()]
		self.total_cost = round(sum(self.total_cost), 2)

		# Tree view of the tests and the costs involved
		tk.Frame(
			self.master,
			width = 780,
			height = 247,
			background = 'white',
			highlightthickness = 4).place(x = 10, y = 303)
		self.style = ttk.Style(self.master)
		self.style.configure('Treeview', font = (None, 10), rowheight = 21)
		self.test_cost_display = ttk.Treeview(self.master)
		self.scroll_bar = ttk.Scrollbar(
			self.master,
			orient = 'vertical',
			command = self.test_cost_display.yview)
		self.scroll_bar.place(x = 767, y = 309, height = 235)
		self.test_cost_display.configure(yscrollcommand = self.scroll_bar.set)
		self.test_cost_display["columns"] = ("#1", "#2")
		self.test_cost_display.column(
			"#0", width = 50, minwidth = 70, stretch = tk.NO)
		self.test_cost_display.heading(
			"#0", text = "S.No.", anchor = tk.CENTER)
		self.test_cost_display.column(
			"#1", width = 600, minwidth = 70, stretch = tk.NO)
		self.test_cost_display.heading(
			"#1", text = "Test Details", anchor = tk.CENTER)
		self.test_cost_display.column(
			"#2", width = 117, minwidth = 77, stretch = tk.NO, anchor = tk.CENTER)
		self.test_cost_display.heading(
			"#2", text = "Cost (EUR)", anchor = tk.CENTER)
		for index, data in enumerate(self.test_cost, start = 1):
			self.test_cost_display.insert(
				"",
				index,
				text = str(index),
				values = (data[0], str(data[1])))
		self.test_cost_display.place(x = 15, y = 308)

		# Total cost
		f_cost = format_currency(self.total_cost, 'EUR', locale='de_DE')

		tk.Label(
			self.master, text = "Total: ",
			bg = 'white', fg = 'black',
			font = ('arial bold', 13)).place(x = 610, y = 555)
		tk.Label(
			self.master, text = str(f_cost),
			bg = 'white', fg = 'green',
			font = ('helvetica', 12, 'bold')).place(x = 660, y = 555)

		ttk.Button(
			self.master,
			text = "Cancel",
			command = self.master.destroy).place(x = 710, y = 600)
		ttk.Button(
			self.master,
			text = "Confirm",
			command = lambda: self.generate_pdf(
				self.test_,
				self.cost_,
				**self.input_values)).place(x = 610, y = 600)

	def generate_pdf(self, test, cost, **kwargs):
		"""
		Generates pdf file based on user request

		Parameters:
		----------
			test : dict
				Collection of work package ids and test names

			costs : dict
				Collection of work package ids and cost names

			**kwargs : dict
				Contains change type, subassembly, part name, 
				requester, creator and comment values

		Return:
		------
			None
		"""

		# Confirmation from the user for generating template
		self.confirm_message = "Are you sure ?"
		self.confirm_ = messagebox.askyesno(
			"Confirmation", 
			self.confirm_message)

		if self.confirm_:
			hdp_data = TransmissionTemplate(
				test, cost, **kwargs)
			hdp_data.generate_template(**records)
			self.master.destroy()
		else:
			self.master.destroy()
			return

class Settings:
	"""
	A class to represent the settings window of the application

	Attributes:
	----------
		master : tkinter.Tk class
			Base class for the construction of the settings window

	Methods:
	------
		report : Generates the usage information of the application

		update_test_path : For updating the test database path

		update_cost_path : For updating the cost database path

		save_paths : Saves the path selected by the user

		update_database : Updates the path in the SQL database

		data_addition : For adding new requesters and creators

		data_deletion : For removing existing requesters and creators

		name_shortner : Limiting the file name to less than 50 chars
	"""

	def __init__(self, master):
		"""
		Constructs the settings window of the application

		Parameters:
		----------
			master : tkinter.Tk class
				Base class for the construction of the settings window
		"""

		# Basic configuration of the window
		self.master = master
		self.master.title(SETTINGS_WINDOW_TITLE)
		self.master.geometry(SETTINGS_WINDOW_RESOLUTION)
		self.master.resizable(0, 0)
		self.master.configure(background = 'white')

		# Assigning required identifiers
		self.initial_requesters = "\n".join(requesters)
		self.initial_creators = "\n".join(creators)
		self.new_test_path = str()
		self.new_cost_path = str()

		# Shortening the test and cost database names
		self.test_db_name = Settings.name_shortner(
			os.path.basename(databases["Test"]))
		self.cost_db_name = Settings.name_shortner(
			os.path.basename(databases["Cost"]))

		# ---------------------------DATABASES-------------------------------

		tk.Frame(
			self.master,
			width = 540,
			height = 40,
			background = 'dark red',
			highlightthickness = 4).pack()
		tk.Label(
			self.master, text = "Databases",
			fg = 'white', bg = 'dark red',
			font = ('helvetica 14 bold')).place(x = 10, y = 5)

		# Test database
		tk.Frame(
			self.master,
			width = 540,
			height = 40,
			background = 'white',
			highlightthickness = 4).pack()
		tk.Label(
			self.master, text = "TEST : ",
			fg = 'black', bg = 'white',
			font = ('helvetica 12 bold')).place(x = 10, y = 49)
		tk.Label(
			self.master, text = self.test_db_name,
			fg = 'dark green', bg = 'white',
			font = ('helvetica 10')).place(x = 70, y = 50)
		ttk.Button(
			self.master,
			text = 'Change',
			command = self.update_test_path).place(x = 460, y = 48)

		# Cost database
		tk.Frame(
			self.master,
			width = 540,
			height = 40,
			background = 'white',
			highlightthickness = 4).pack()
		tk.Label(
			self.master, text = "COST : ",
			fg = 'black', bg = 'white',
			font = ('helvetica 12 bold')).place(x = 10, y = 88)
		tk.Label(
			self.master, text = self.cost_db_name,
			fg = 'dark green', bg = 'white',
			font = ('helvetica 10')).place(x = 70, y = 89)
		ttk.Button(
			self.master,
			text = 'Change',
			command = self.update_cost_path).place(x = 460, y = 88)

		# ---------------------REQUESTERS AND CREATORS------------------------

		tk.Frame(
			self.master,
			width = 540,
			height = 40,
			background = 'dark red',
			highlightthickness = 4).pack()
		tk.Label(
			self.master, text = "Requesters & Creators",
			fg = 'white', bg = 'dark red',
			font = ('helvetica 14 bold')).place(x = 10, y = 125)
		tk.Frame(
			self.master,
			width = 540,
			height = 150,
			background = 'white',
			highlightthickness = 4).pack()

		# Requesters
		self.requesters = tkst.ScrolledText(
			master = self.master,
			wrap = tk.WORD,
			width = 30,
			height = 8)
		self.requesters.insert(
			tk.INSERT,
			self.initial_requesters.strip())
		self.requesters.config(state = tk.DISABLED)
		self.requesters.place(x = 12, y = 170)

		# Creators
		self.creators = tkst.ScrolledText(
			master = self.master,
			wrap = tk.WORD,
			width = 30,
			height = 8)
		self.creators.insert(
			tk.INSERT,
			self.initial_creators.strip())
		self.creators.config(state = tk.DISABLED)
		self.creators.place(x = 278, y = 170)

		# Add and remove options
		tk.Frame(
			self.master,
			width = 540,
			height = 40,
			background = 'white',
			highlightthickness = 4).pack()
		ttk.Button(
			self.master,
			text = "Add",
			command = self.data_addition).place(x = 180, y = 318)
		ttk.Button(
			self.master,
			text = "Remove",
			command = self.data_deletion).place(x = 280, y = 318)

		# ------------------------------REPORT-------------------------------

		self.report_msg = "Generate usage information of the application"

		tk.Frame(
			self.master,
			width = 540,
			height = 40,
			background = 'dark red',
			highlightthickness = 4).pack()
		tk.Frame(
			self.master,
			width = 540,
			height = 60,
			background = 'white',
			highlightthickness = 4).pack()
		tk.Label(
			self.master, text = "Report",
			fg = 'white', bg = 'dark red',
			font = ('helvetica 14 bold')).place(x = 10, y = 355)
		tk.Label(
			self.master, text = self.report_msg,
			fg = 'black', bg = 'white',
			font = ('helvetica 10 italic')).place(x = 20, y = 408)
		ttk.Button(
			self.master,
			text = "Report",
			command = self.report).place(x = 450, y = 408)


		ttk.Button(
			self.master,
			text = "Save",
			command = self.save_paths).place(x = 360, y = 460)

		ttk.Button(
			self.master,
			text = "Cancel",
			command = self.master.destroy).place(x = 460, y = 460)

	def report(self):
		"""
		Generates the usage information of the application

		Parameters:
		----------
			None

		Return:
		------
			None
		"""

		try:
			report_data = TransmissionReport(records["Report"])
			report_data.generate_report()
			messagebox.showinfo("Success", "The report has been generated")
		except Exception as e:
			messagebox.showwarning(
				"Report Error",
				"Sorry..! Could not generate report. "+str(e))
			logging.error(traceback.format_exc())
		self.master.destroy()

	def update_test_path(self):
		"""
		Fetches the new path for the test database and updates

		the settings window

		Parameters:
		----------
			None

		Return:
		------
			None
		"""

		self.new_path = askopenfile()
		self.new_test_path = self.new_path.name
		self.test_basefile = os.path.basename(self.new_test_path)
		self.test_basefile = Settings.name_shortner(self.test_basefile)
		tk.Label(
			self.master, text = self.test_basefile,
			fg = 'dark blue', bg = 'white',
			font = ('helvetica 10')).place(x = 70, y = 50)


	def update_cost_path(self):
		"""
		Fetches the new path for the cost database and updates

		the settings window

		Parameters:
		----------
			None

		Return:
		------
			None
		"""

		self.new_path = askopenfile()
		self.new_cost_path = self.new_path.name
		self.cost_basefile = os.path.basename(self.new_cost_path)
		self.cost_basefile = Settings.name_shortner(self.cost_basefile)

		tk.Label(
			self.master, text = self.cost_basefile,
			fg = 'dark blue', bg = 'white',
			font = ('helvetica 10')).place(x = 70, y = 89)

	def save_paths(self):
		"""
		Saves the new path for the test and cost database

		if they are modified

		Parameters:
		----------
			None

		Return:
		------
			None
		"""

		result = messagebox.askyesno("Confirmation", "Are you sure?")
		if result:
			if self.new_test_path:
				self.update_database(self.new_test_path, "Test")
			if self.new_cost_path:
				self.update_database(self.new_cost_path, "Cost")
			if self.new_test_path or self.new_cost_path:
				messagebox.showinfo(
					"Success", 
					"""The database has been successfully updated
					\t\tKindly restart the application""")
				sys.exit(0)
		else:
			pass

	def update_database(self, path, field):
		"""
		Updates the SQL database with the provided path under the 

		test and cost path fields

		Parameters:
		----------
			path : str
				New path selected by the user

			field : str
				Specific field of the path provided

		Return:
		------
			None
		"""

		try:
			database = sqlite3.connect('database/info.db')
			cur = database.cursor()
		except Exception as e:
			messagebox.showwarning("Database Error", str(e))
			logging.error(traceback.format_exc())
			sys.exit(0)
		else:
			cur.execute('''UPDATE Databases 
				SET Path = ? WHERE Name = ?''',(path, field))
			database.commit()
		finally:
			database.close()

	def data_addition(self):
		"""
		Instantiate the data addition window inheriting from the

		tkinter Toplevel class

		Parameters:
		----------
			None

		Return:
		------
			None
		"""

		self.data_addition_window = tk.Toplevel(self.master)
		self.data_addition_app = DataAddition(self.data_addition_window)

	def data_deletion(self):
		"""
		Instantiate the data deletion window inheriting from the

		tkinter Toplevel class

		Parameters:
		----------
			None

		Return:
			None
		"""

		self.data_deletion_window = tk.Toplevel(self.master)
		self.data_deletion_app = DataDeletion(self.data_deletion_window)

	@staticmethod
	def name_shortner(name):
		"""
		Shortens the name by limiting the characters to less than 50

		Parameters:
		----------
			name : str
				The name of the test or cost database

		Return:
		------
			None
		"""

		if len(name) > 50:
			return name[:20] + "." * 5 + name[-20:]
		else:
			return name

class DataAddition:
	"""
	A class for representing the data addition window of the application

	Attributes:
	----------
		master : tkinter.Tk class
			Base class for construction of the data addition window

	Method:
		add_data : Fetches the new requester/creator name

		update_database : Updates the SQL database with the new data
	"""
	
	def __init__(self, master):
		"""
		Constructs the data addition window of the application

		Parameters:
		----------
			master : tkinter.Tk class
				Base class for the constructin of the window
		"""

		# Basic configuration of the window
		self.master = master
		self.master.title(DATA_ADDITION_WINDOW_TITLE)
		self.master.geometry(DATA_ADDITION_WINDOW_RESOLUTION)
		self.master.resizable(0, 0)
		self.master.configure(background = 'white')

		# Assigning required input variable wrappers
		self.selections = tk.StringVar()

		tk.Label(
			self.master, text = "Category:",
			fg = 'black', bg = 'white',
			font = ('helvetica 12 bold')).place(x = 20, y = 5)
		ttk.Combobox(
			self.master, values = ["Requesters", "Creators"],
			width = 20, foreground = 'dark blue',
			state = "readonly",
			textvariable = self.selections).place(x = 110, y = 7.5)
		tk.Label(
			self.master, text = "Name:",
			fg = 'black', bg = 'white',
			font = ('helvetica 12 bold')).place(x = 20, y = 40)
		self.text = tk.Text(
			self.master,
			width = 20,
			bd = 1.5,
			height = 1)
		self.text.place(x = 110, y = 42.5)
		ttk.Button(
			self.master,
			text = "Save",
			command = lambda: self.add_data(
				self.selections, self.text)).place(x = 60, y = 85)
		ttk.Button(
			self.master,
			text = "Cancel",
			command = self.master.destroy).place(x = 160, y = 85)

	def add_data(self, value, name):
		"""
		Validates the inputs and triggers the updation based on user

		confirmation

		Parameters:
		----------
			value : tkinter.StringVar
				Variable wrapper for capturing the requester 
				or creator category

			name : str
				New user information entered in the application

		Return:
		------
			None
		"""

		self.table = value.get()
		self.name = name.get("1.0", "end-1c")
		confirmation = messagebox.askyesno("Confirmation", "Are you sure?")
		if confirmation:
			if self.table and self.name:
				self.update_database()
			else:
				messagebox.showwarning("Error", "Insufficient inputs")
				self.master.destroy()
		else:
			pass

	def update_database(self):
		"""
		Updates the SQL database with the new creator or requester

		Parameters:
		----------
			None

		Return:
		------
			None
		"""

		try:
			database = sqlite3.connect('database/info.db')
			cur = database.cursor()
		except Exception as e:
			messagebox.showwarning("Database Error", str(e))
			logging.error(traceback.format_exc())
			sys.exit(0)
		else:
			cur.execute(
				'''INSERT INTO 
				%s(Name)VALUES(?)'''%self.table,(self.name,))
			database.commit()
			messagebox.showinfo(
				"Success", 
				"""The database is updated successfully. 
					\t\tKindly restart the application to view the changes""")
		finally:
			database.close()
			sys.exit(0)

class DataDeletion:
	"""
	A class for representing the data deletion window of the application

	Attributes:
	----------
		master : tkinter.Tk class
			Base class for the construction of the data deletion window

	Method:
	------
		remove_data : Removes the selected requester or creator

		update_database : Updates the SQL database with the changes
	"""

	def __init__(self, master):
		"""
		Constructs the data deletion window of the application

		Parameters:
		----------
			master : tkinter.Tk class
				Base class for the construction of the window
		"""

		# Basic configuration of the window
		self.master = master
		self.master.title(DATA_DELETION_WINDOW_TITLE)
		self.master.geometry(DATA_DELETION_WINDOW_RESOLUTION)
		self.master.resizable(0, 0)
		self.master.configure(background = 'white')

		# Assigning required input variable wrappers
		self.selection = tk.StringVar()

		# Requesters
		tk.Label(
			self.master, text = "Requesters",
			fg = 'black', bg = 'white',
			font = ('helvetica 12 bold')).place(x = 12, y = 5)
		self.requesterbox = tk.Listbox(
			self.master, bg = 'white',
			width = 30, height = 8,
			cursor = 'hand2')
		self.requester_sr = tk.Scrollbar(self.master)
		self.requesterbox.config(yscrollcommand = self.requester_sr.set)
		self.requester_sr.config(command = self.requesterbox.yview)
		self.requester_sr.place(x = 200, y = 30, height = 130)
		for colleague in requesters:
			self.requesterbox.insert(tk.END, colleague)
		self.requesterbox.place(x = 12, y = 30)

		# Creators
		tk.Label(
			self.master, text = "Creators",
			fg = 'black', bg = 'white',
			font = ('helvetica 12 bold')).place(x = 232, y = 5)
		self.creatorbox = tk.Listbox(
			self.master, bg = 'white',
			width = 30, height = 8,
			cursor = 'hand2')
		self.creator_sr = tk.Scrollbar(self.master)
		self.creatorbox.config(yscrollcommand = self.creator_sr.set)
		self.creator_sr.config(command = self.creatorbox.yview)
		self.creator_sr.place(x = 420, y = 30, height = 130)
		for colleague in creators:
			self.creatorbox.insert(tk.END, colleague)
		self.creatorbox.place(x = 232, y = 30)

		# Category selection
		tk.Label(
			self.master, text = "Select the category:",
			fg = 'black', bg = 'white',
			font = ('helvetica 12 bold')).place(x = 10, y = 170)
		ttk.Combobox(
			self.master, values = ["Requesters", "Creators"],
			width = 20, foreground = 'dark blue',
			state = "readonly",
			textvariable = self.selection).place(x = 200, y = 172.5)

		ttk.Button(
			self.master,
			text = "Remove",
			command = self.remove_data).place(x = 250, y = 210)
		ttk.Button(
			self.master,
			text = "Cancel",
			command = self.master.destroy).place(x = 350, y = 210)

	def remove_data(self):
		"""
		Removes the selected requester or creator

		Parameters:
		----------
			None

		Return:
		------
			None
		"""

		self.category = self.selection.get()
		self.requester_ = self.requesterbox.get(tk.ACTIVE)
		self.creator_ = self.creatorbox.get(tk.ACTIVE)
		result = messagebox.askyesno(
			"Confirmation",
			"Are you sure to remove from %s?"%self.category)
		if result:
			if self.category == "Requesters":
				print("requester", self.requester_)
				self.update_database(self.requester_)
			elif self.category == "Creators":
				print("creator", self.creator_)
				self.update_database(self.creator_)
			else:
				messagebox.showwarning(
					"Error", 
					"Please select a category")
				self.master.destroy()
		else:
			self.master.destroy()

	def update_database(self, value):
		"""
		Updates the SQL database by removing the selected 

		requester or creator

		Parameters:
		----------
			None

		Return:
		------
			None
		"""
		
		try:
			database = sqlite3.connect('database/info.db')
			cur = database.cursor()
			cur.execute('''DELETE FROM %s 
				WHERE Name = ?'''%self.category, (value,))
			database.commit()
		except Exception as e:
			messagebox.showwarning("Updation Error", str(e))
			logging.error(traceback.format_exc())
			sys.exit(0)
		else:
			messagebox.showinfo(
				"Success", 
				"""The database has been updated successfully
					\t\tKindly restart the application to view the changes""")
		finally:
			database.close()
			sys.exit(0)

if __name__ == "__main__":
	executor = ThreadPoolExecutor(max_workers = 2)
	test_database = executor.submit(MainWindow.load_databases, databases["Test"])
	cost_database = executor.submit(MainWindow.load_databases, databases["Cost"])
	window = tk.Tk()
	application = MainWindow(window)
	window.mainloop()