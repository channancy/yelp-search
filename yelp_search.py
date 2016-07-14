from yelp.client import Client
from yelp.oauth1_authenticator import Oauth1Authenticator
import io, json, openpyxl, sys

# Read API keys
with io.open('config_secret.json') as cred:
    creds = json.load(cred)
    auth = Oauth1Authenticator(**creds)
    client = Client(auth)

# Open Excel document for writing
wb = openpyxl.Workbook()
dest_filename = 'search_results.xlsx'
sheet = wb.get_active_sheet()

# Write labels on first row
rowNum = 1
colNum = 1

sheet.cell(row=rowNum, column=colNum).value = 'Search'
sheet.cell(row=rowNum, column=colNum + 1).value = 'Name'
sheet.cell(row=rowNum, column=colNum + 2).value = 'Rating'
sheet.cell(row=rowNum, column=colNum + 3).value = 'Number of Reviews'
sheet.cell(row=rowNum, column=colNum + 4).value = 'Street Address'
sheet.cell(row=rowNum, column=colNum + 5).value = 'City'
sheet.cell(row=rowNum, column=colNum + 6).value = 'State'
sheet.cell(row=rowNum, column=colNum + 7).value = 'Zip Code'
sheet.cell(row=rowNum, column=colNum + 8).value = 'Phone Number'
sheet.cell(row=rowNum, column=colNum + 9).value = 'Yelp Link'

# Write data starting on second row
rowNum = 2

# Check command line arguments
if len(sys.argv) < 3:
	print 'Usage: python ' + sys.argv[0] + ' <filename> <location>'
	exit(1)

# Filename and location are passed in
filename = sys.argv[1]
location = sys.argv[2]

# Open file for reading
with open(filename) as f:

	# Each line is a separate search
	for line in f:

		# Start again at first column
		colNum = 1

		# Strip whitespace
		line_search = line.strip()

		# Search parameters
		params = {
		    'term': line_search,
		}

		# Yelp Search API
		response = client.search(location, **params)

		try:
			# First result
			business = response.businesses[0]

			name = business.name
			rating = business.rating
			review_count = business.review_count
			address_list = business.location.address
			# http://stackoverflow.com/questions/18272066/easy-way-to-convert-a-unicode-list-to-a-list-containing-python-strings
			address_encoded = [x.encode('UTF8') for x in address_list]
			# http://stackoverflow.com/questions/5618878/how-to-convert-list-to-string
			address = ' '.join(address_encoded)
			city = business.location.city
			state = business.location.state_code
			zipcode = business.location.postal_code
			city_state_zipcode = city + ', ' + state + ' ' + zipcode
			phone = business.display_phone
			url = business.url

			# Display search results
			print line_search
			print name
			print rating
			print 'Number of reviews: ' + str(review_count)
			print address
			print city_state_zipcode
			print phone
			print url
			print ''

			# Write search results to Excel document
			sheet.cell(row=rowNum, column=colNum).value = line_search
			# Each business occupies one row, information across columns
			colNum += 1

			sheet.cell(row=rowNum, column=colNum).value = name
			colNum += 1

			sheet.cell(row=rowNum, column=colNum).value = rating
			colNum += 1

			sheet.cell(row=rowNum, column=colNum).value = review_count
			colNum += 1

			sheet.cell(row=rowNum, column=colNum).value = address
			colNum += 1

			sheet.cell(row=rowNum, column=colNum).value = city
			colNum += 1

			sheet.cell(row=rowNum, column=colNum).value = state
			colNum += 1

			sheet.cell(row=rowNum, column=colNum).value = zipcode
			colNum += 1

			sheet.cell(row=rowNum, column=colNum).value = phone
			colNum += 1

			sheet.cell(row=rowNum, column=colNum).value = url
			# End of business, move to next row
			rowNum += 1

		# If list index out of range because no search results found
		# Simply print/write search
		except IndexError:
			print line_search
			print ''

			sheet.cell(row=rowNum, column=colNum).value = line_search
			rowNum += 1

		# If NoneType because of missing information
		# Simply print/write search
		except TypeError:
			print line_search
			print ''

			sheet.cell(row=rowNum, column=colNum).value = line_search
			rowNum += 1

		# If user hits Ctrl+C
		except KeyboardInterrupt:
			print 'Hit Ctrl+C'
			sys.exit()

		finally:
			# Save Excel document before raise exceptions
			wb.save(filename=dest_filename)

# Save Excel document
wb.save(filename=dest_filename)