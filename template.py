"""
A template module for generating the Test and Cost Template for
the Sample Product
"""

__author__ = "Monish Mohanan"
__version__ = "1.0"

# Importing required libraries
try:
	import getpass
	import os
	import sqlite3
	import decimal
	from datetime import datetime
	from fpdf import FPDF
	from tkinter import messagebox
	from babel.numbers import format_currency
except ImportError as e:
	from tkinter import messagebox
	messagebox.showwarning("Import Error", str(e))

# Defining the necessary constants
BOSCH_LOGO_IMAGE = "images/logo.png"
TEMPLATE_FOLDER = "Templates/"
TITLE = "TRANS | TEST AND COST TEMPLATE"
DETAILS = "REQUEST FOR DEV-COST"
TEST_AND_COST = "TEST TO BE PERFORMED"
REMARKS = "GENERAL REMARKS"
COMMENTS = "USER COMMENTS"
NAME = "transmission_test_cost_template"
LIMIT = 77

class TransmissionTemplate:
	"""
	A class for generating Test & Cost template

	Attributes:
	----------
		tests : dict
			Collection of work package ids & test names 

		costs : dict
			Collection of work package ids and cost names

		**kwargs : dict
			Contains change type, subassembly, part, requester,
			creator, comment values

	Method:
	------
		generate_template : Generates the test and cost template
	"""

	def __init__(self, tests, costs, **kwargs):
		"""
		Constructs the required identifier for generating the 
		template

		Parameters:
		----------
			tests : dict
				Collection of work package ids and test names

			costs : dict
				Collection of work package ids and cost names

			**kwargs : dict
				Contains change type, subassembly, part name,
				requester, creator and comment values
		"""

		# Assigning the identifiers for template generation
		self.test_ = tests
		self.cost_ = costs
		self.input_values = kwargs

		# Assigning the document name
		self.user = getpass.getuser().lower()
		self.name = "_".join((self.user, NAME))

		# Formatting the change type for the template
		self.format_changetype = self.input_values["Change Type"]
		self.split_words = self.format_changetype.split()
		if self.split_words[3] == "@":
			self.format_changetype = " ".join(self.split_words[:3])
		else:
			self.format_changetype = " ".join(self.split_words[:4])
		self.input_values["Change Type"] = self.format_changetype

		# Assigning general input fields
		self.detail_one = {
			"Reason" : "Sample",
			"Details" : "None",
			"Type of change" : self.input_values["Change Type"],
			"Requester (Name, Dept.)" : self.input_values["Requester"],
			"Platform:" : "Transmission",
			"Subassembly" : self.input_values["Subassembly"],
			"Part Name" : self.input_values["Part Name"]
			}

		created_by = self.input_values["Creator"].split(" ")
		name_and_dept = "\n".join((" ".join(created_by[:-1]),created_by[-1]))
		self.detail_two = {
			"Created By" : name_and_dept,
			"Checked By" : " \n ",
			"Approved By" : " \n "
			}

		# Assigning column headers and its spacing
		self.column_headers = {
			"S.No." : 10,
			"Test Name" : 125,
			"Cost" : 25,
			"Remarks" : 30
			}

		# Assigning padding values for the data
		self.padding = dict()
		for test in self.test_.values():
			if len(test) > LIMIT:
				self.padding[test] = 2
			else:
				self.padding[test] = 1

		# Assigning page width and default cell height
		self.width = 190
		self.height = 8

		self.total_test = len(self.test_)
		self.total_cost = str(round(sum(self.cost_.values()), 2))
		self.total = format_currency(self.total_cost, 'EUR', locale = 'de_DE')[:-2]
		self.total = str(self.total).split(',')[0]

		try:
			data = sqlite3.connect('database/info.db')
			cur = data.cursor()
		except Exception as e:
			messagebox.showwarning("Database Error", str(e))
		else:
			cur.execute('''SELECT Name, Path FROM Databases''')
			databases = dict(
				(col[0], col[1]) for col in cur.fetchall())
			self.test_name = os.path.basename(databases['Test'])
			self.cost_name = os.path.basename(databases['Cost'])
		finally:
			data.close()

		if not os.path.exists(TEMPLATE_FOLDER):
			os.mkdir(TEMPLATE_FOLDER.split('/')[0])

		
	def generate_template(self, **path):
		"""
		Generate the test and cost template in the present working 
		directory and also a copy in the server. Record the entry
		in the SQL database

		Parameters:
		----------
			**path : dict
				Contains the location of server folder and report

		Return:
		------
			None
		"""

		# Assigning the required identifiers
		self.record_inputs = path
		self.record_folder = path["Records"]
		self.generated = False

		self.now = datetime.now()
		self.date = self.now.strftime("%d/%m/%Y")
		self.time = self.now.strftime("%H:%M:%S")

		# PDF object with A4 sheet size and Portrait orientation
		self.pdf = FPDF(orientation = 'P', unit = 'mm', format = 'A4')
		self.pdf.add_page()
		self.pdf.set_fill_color(220, 220, 220)

		# Logo
		self.pdf.set_font('Arial', 'B', size = 20)
		self.pdf.image(BOSCH_LOGO_IMAGE, x = 11, y = 11, w = 40, h = 8.5)
		self.pdf.cell(42.5, 10.5, border = 1)

		# Title 
		self.pdf.set_font('Arial', 'B', size = 20)
		self.pdf.set_text_color(0, 0, 0)
		self.pdf.cell(147.5, 10.5, txt = TITLE, align = 'C', border = 1)
		self.pdf.ln()

		# First sub heading
		self.pdf.set_font('Arial', 'B', size = 12)
		self.pdf.cell(
			self.width, 
			self.height, 
			txt = DETAILS, 
			border = 1, 
			fill = True)
		self.pdf.ln()

		# General information
		self.x1, self.y1 = self.pdf.get_x(), self.pdf.get_y()
		for key, value in self.detail_one.items():
			self.pdf.set_font('Arial', 'B', size = 10)
			self.pdf.cell(50, 6, txt = key, border = 1)
			self.pdf.set_font('Arial', size = 9)
			self.pdf.cell(62, 6, txt = value, border = 1)
			self.pdf.ln()

		self.pdf.set_font('Arial', 'B', size = 10)
		self.pdf.set_xy(self.x1 + 112, self.y1)
		for key in self.detail_two.keys():
			self.pdf.set_x(self.x1 + 112)
			self.pdf.cell(35, 14, txt = key, border = 1, align = 'C')
			self.pdf.ln()

		self.pdf.set_font('Arial', size = 9)
		self.pdf.set_xy(self.x1 + 147, self.y1)
		for value in self.detail_two.values():
			self.pdf.set_x(self.x1 + 147)
			self.pdf.multi_cell(43, 7, txt = value, border = 1, align = 'C')

		# Second sub heading
		self.pdf.set_font('Arial', 'B', size = 12)
		self.pdf.cell(
			self.width,
			self.height,
			txt = TEST_AND_COST,
			border = 1,
			fill = True)
		self.pdf.ln()

		# Creating columns for the body of the template
		self.pdf.set_font('Arial', 'B', size = 10)
		for key, value in self.column_headers.items():
			self.pdf.cell(value, self.height, txt = key, border = 1, align = 'C')
		self.pdf.ln()

		# Creating the serial numbers 
		self.x2, self.y2 = self.pdf.get_x(), self.pdf.get_y()
		self.pdf.set_font('Arial', size = 9)
		for index, pad in enumerate(self.padding.values(), start = 1):
			self.pdf.cell(
				self.column_headers["S.No."], 
				self.height * pad, 
				txt = str(index), 
				border = 1, 
				align = 'C')
			self.pdf.ln()

		# Writing the test names
		self.pdf.set_xy(self.x2 + 10, self.y2)
		for test in self.test_.values():
			self.pdf.set_x(self.x2 + 10)
			self.pdf.multi_cell(
				self.column_headers["Test Name"], 
				self.height, 
				txt = test, 
				border = 1, 
				align = 'L')

		# Writing the cost values
		self.pdf.set_xy(self.x2 + 135, self.y2)
		for cost, pad in zip(self.cost_.values(), self.padding.values()):
			self.pdf.set_x(self.x2 + 135)
			self.pdf.cell(
				self.column_headers["Cost"],
				self.height * pad,
				txt = str(cost),
				border = 1,
				align = 'C')
			self.pdf.ln()

		# Remark column
		self.pdf.set_xy(self.x2 + 160, self.y2)
		for blank, pad in zip(range(len(self.padding)), self.padding.values()):
			self.pdf.set_x(self.x2 + 160)
			self.pdf.cell(
				self.column_headers["Remarks"],
				self.height * pad,
				txt = " ",
				border = 1)
			self.pdf.ln()

		# Total cost value
		self.pdf.cell(
			self.width - 55, 
			self.height, 
			txt = "TOTAL COST", 
			border = 1, 
			align = 'C')
		self.pdf.cell(
			self.column_headers["Cost"],
			self.height,
			txt = " ".join((self.total_cost, "EUR")),
			border = 1,
			align = 'C')
		self.pdf.cell(
			self.column_headers["Remarks"], 
			self.height,
			txt = " ",
			border = 1)
		self.pdf.ln()

		# General remarks
		self.pdf.set_font('Arial', 'B', size = 12)
		self.pdf.cell(
			self.width,
			self.height,
			txt = REMARKS,
			border = 1, 
			fill = True)
		self.pdf.ln()
		self.pdf.cell(self.width, self.height + 4, txt = " ", border = 1)
		self.pdf.ln()

		# User comments
		self.pdf.cell(
			self.width,
			self.height,
			txt = COMMENTS,
			border = 1,
			fill = True)
		self.pdf.ln()
		self.pdf.set_font('Arial', size = 9)
		self.pdf.cell(
			self.width,
			self.height + 4,
			txt = self.input_values["Comment"],
			border = 1,
			align = 'C')

		# Footer
		self.pdf.ln()
		self.pdf.set_font('Arial', size = 6)
		self.pdf.cell(
						self.width/2,
						self.height/2,
						txt = "Created based on: "+ str(self.test_name),
						border = 1,
						align = 'L')
		self.pdf.cell(
						self.width/2,
						self.height/2,
						txt = "Cost based on: "+ str(self.cost_name),
						border = 1,
						align = 'L')

		# Creating PDFs and recording the entry in teh SQL database
		try:
			self.report = sqlite3.connect(self.record_inputs["Report"])
			self.cur = self.report.cursor()
		except Exception as e:
			messagebox.showwarning(
				"Report Error",
				"This record will not be captured"+str(e))
		else:
			self.cur.execute('''SELECT ID FROM Record''')
			self.ids = self.cur.fetchall()
			self.new_id = len(self.ids) + 1
			self.pdf_name = "_".join((self.name, str(self.new_id)))
			self.pdf_name = ".".join((self.pdf_name, "pdf"))
			self.pdf.output(TEMPLATE_FOLDER + self.pdf_name)
			self.path = str(os.path.join(self.record_folder, self.pdf_name))
			self.pdf.output(self.path)
			self.cur.execute('''INSERT INTO Record(Date, Time, Requester,Creator, 
				Changetype, Test, Cost, Link, User, Subassembly, Partname
				)VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',(
					self.date, self.time, self.input_values["Requester"],
					self.input_values["Creator"], self.input_values["Change Type"],
					self.total_test, self.total_cost, self.path, self.user, 
					self.input_values["Subassembly"], self.input_values["Part Name"]))
			self.report.commit()
			self.generated = True
		finally:
			self.report.close()

		# Check if the PDFs are generated and the records are updated
		# Display the output message to the user
		if self.generated:
			messagebox.showinfo(
				"Success", 
				"The Test Cost information is generated")
		else:
			messagebox.showwarning(
				"Failure", 
				"The Test Cost information could not be generated")


