from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.shared import Pt
import csv

def write_line(cell, attr, font_size = 25):
	"""Writes info to the DOCX cell"""

	paragraph = cell.add_paragraph()
	p_format = paragraph.paragraph_format

	# no vertical spacing
	p_format.space_after = Pt(0)

	# centering
	p_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

	# writing
	run = paragraph.add_run(attr)

	# font size
	font = run.font
	font.size = Pt(font_size)

def stickers_transform():
	"""Transforms the CSV data into a DOCX file"""
	document = Document('landscape.docx')

	# for section in document.sections:
	# 	section.orientation = WD_ORIENT.LANDSCAPE

	with open('../orders_export 4.csv', 'rU') as infile, open('../bad_orders.csv','w') as outfile:
		reader = csv.reader(infile)
		reader.next()

		writer = csv.writer(outfile)

		table = document.add_table(2,2)

		i = 0
		for row in reader:
			if row[2]:
				index1, index2 = i / 2, i % 2
				cell = table.cell(index1, index2)
				
				# product name
				prod = row[17].lower()

				if 'refrigerator' in prod:
					prod = 'FRIDGE'
				elif 'room ready' in prod:
					prod = prod.split('-')[-1].strip().upper().split('/')[0]
				else:
					print 'Unavailable product: ', prod
					continue

				# personal attributes
				attrs = row[45]
				attrs = attrs.upper().split('\n')

				try:
					# no need for email, reference, or netid
					attrs = [attr for attr in attrs
							if ('EMAIL' not in attr) and ('REFERENCE' not in attr) and ('NETID' not in attr)]

					name = attrs[0][9:].split()
					attrs[0] = name[0] + ' ' + name[-1]
					attrs[1] = attrs[1][10:].split('(')[-1].rstrip(')') # dorm
					attrs[2] = attrs[2][12:] # room

					# just switching up order
					attrs[0], attrs[1], attrs[2] = attrs[1], attrs[2], attrs[0]

					write_line(cell, '\n', 20)
					for attr in attrs:
						write_line(cell, attr)

				except:
					print 'Information error: ', row[0]
					writer.writerow(row)
					continue

				write_line(cell, prod)
				write_line(cell, '\n', 20)

				i += 1

			if i % 4 == 0:
				table = document.add_table(2,2)
				i = 0

	document.save('../labels.docx')

if __name__ == '__main__':
	stickers_transform()