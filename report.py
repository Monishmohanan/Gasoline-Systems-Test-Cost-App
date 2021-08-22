"""
A report module for generating the usage information of the Test
and Cost Template application
"""

__author__ = "Monish Mohanan"
__version__ = "1.0"

# Importing required libraries
try:
	import os
	import sys
	import sqlite3
	from fpdf import FPDF
	from tkinter import messagebox
	from babel.numbers import format_currency
except ImportError as e:
	from tkinter import messagebox
	messagebox.showwarning("Import Error", str(e))
	sys.exit(0)

class TransmissionReport:
	"""
	A class for generating the Test & Cost template 

	usage information

	Attributes:
	----------
		path : str
			Location of the report database in the server

	Method:
	------
		generate_report : Generates the usage information
	"""

	def __init__(self, path):
		"""
		Constructs the required identifiers for generating the

		usage information as a report

		Parameters:
		----------
			path : str
				Location of the report database in the server
		"""

		# Assigning the identifiers for report generation
		self.title = "Test & Cost Template - Report"
		self.headings = {
							"S.No.":10,
							"Date":30,
							"Requester":45,
							"Creator":45,
							"Change Type":45,
							"Subassembly":45,
							"Cost":20,
							"File":20
						}

		# Querying the SQL database and extracting the information
		try:
			report = sqlite3.connect(path)
			cur = report.cursor()
		except Exception as e:
			messagebox.showwarning("Database Error", str(e))
			sys.exit(0)
		else:
			cur.execute('''SELECT ID FROM Record''')
			self.sno = [col[0] for col in cur.fetchall()]
			cur.execute('''SELECT Date FROM Record''')
			self.date = [col[0] for col in cur.fetchall()]
			cur.execute('''SELECT Requester FROM Record''')
			self.requester = [col[0] for col in cur.fetchall()]
			cur.execute('''SELECT Creator FROM Record''')
			self.creator = [col[0] for col in cur.fetchall()]
			cur.execute('''SELECT Changetype FROM Record''')
			self.changetype = [col[0] for col in cur.fetchall()]
			cur.execute('''SELECT Cost FROM Record''')
			self.cost = [col[0] for col in cur.fetchall()]
			cur.execute('''SELECT Link FROM Record''')
			self.link = [col[0] for col in cur.fetchall()]
			cur.execute('''SELECT User FROM Record''')
			self.user = [col[0] for col in cur.fetchall()]
			cur.execute('''SELECT Subassembly FROM Record''')
			self.subassembly = [col[0] for col in cur.fetchall()]
		finally:
			report.close()

		if not os.path.exists("Report"):
			os.mkdir("Report")

	def generate_report(self):
		"""
		Generates the usage information of the application as a

		report (PDF format) in the present working directory

		Parameters:
		----------
			None

		Return:
			None
		"""

		# PDF object with A4 sheet size and Portrait orientation
		self.pdf = FPDF(orientation = 'L', unit = 'mm', format = 'A4')
		self.pdf.add_page()

		# Title
		self.pdf.set_font("Arial", "B", size = 12)
		self.pdf.set_text_color(255, 255, 255)
		self.pdf.cell(260, 10, txt = self.title, align = 'C', fill = True)
		self.pdf.ln()

		# Headings
		self.pdf.set_font("Arial", "B", size = 10)
		self.pdf.set_text_color(0, 0, 0)
		for key, value in self.headings.items():
			self.pdf.cell(value, 10, txt = key, align = 'C', border = 1)
		self.pdf.ln()

		# Usage information
		self.pdf.set_font("Arial", size = 7)
		for sno, date, requester, creator, cost, link, changetype, \
		subassembly in zip(self.sno, self.date, self.requester, self.creator,
			self.cost, self.link, self.changetype, self.subassembly):
			self.pdf.cell(10, 10, txt = str(sno), align = 'C', border = 1)
			self.pdf.cell(30, 10, txt = date, align = 'C', border = 1)
			self.pdf.cell(45, 10, txt = requester, align = 'C', border = 1)
			self.pdf.cell(45, 10, txt = creator, align = 'C', border = 1)
			self.pdf.cell(45, 10, txt = changetype, align = 'C', border = 1)
			self.pdf.cell(45, 10, txt = subassembly, align = 'C', border = 1)
			self.pdf.cell(20, 10, txt = str(cost), align = 'C', border = 1)
			self.pdf.set_text_color(0, 0, 255)
			self.pdf.cell(
				20, 10, 
				txt = "link", align = 'C', 
				border = 1, link = os.path.join(link.replace("\\", "/")))
			self.pdf.set_text_color(0, 0, 0)
			self.pdf.ln()

		self.pdf.output("report/Test Cost App Usage Report.pdf")

