from yelp.client import Client
from yelp.oauth1_authenticator import Oauth1Authenticator
import io, json, openpyxl, sys

# Read API keys
with io.open('config_secret.json') as cred:
    creds = json.load(cred)
    auth = Oauth1Authenticator(**creds)
    client = Client(auth)

# Desired yelp rating
target = 4

# Open Excel document for writing
wb = openpyxl.load_workbook('search_results.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')

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

# Open text file for writing
# out = open('search_results.txt', 'w')

# Open file for reading
filename = 'Regal Medical Group - PCPs.txt'
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
		response = client.search('Los Angeles', **params)

		try:
			# Print search results with rating of at least target
			# if response.businesses[0].rating >= target:
			name = response.businesses[0].name
			rating = response.businesses[0].rating
			review_count = response.businesses[0].review_count
			address_list = response.businesses[0].location.address
			# http://stackoverflow.com/questions/18272066/easy-way-to-convert-a-unicode-list-to-a-list-containing-python-strings
			address_encoded = [x.encode('UTF8') for x in address_list]
			# http://stackoverflow.com/questions/5618878/how-to-convert-list-to-string
			address = ' '.join(address_encoded)
			city = response.businesses[0].location.city
			state = response.businesses[0].location.state_code
			zipcode = response.businesses[0].location.postal_code
			city_state_zipcode = city + ', ' + state + ' ' + zipcode
			phone = response.businesses[0].display_phone
			url = response.businesses[0].url

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
			rowNum += 1

			# Write search results to text file
			# out.write(line_search + '\n')
			# out.write(name + '\n')
			# out.write(str(rating) + '\n')
			# out.write(str(review_count) + '\n')

		# If list index out of range because no search results found
		except IndexError:
			print line_search
			print ''

			sheet.cell(row=rowNum, column=colNum).value = line_search
			rowNum += 1

		# If NoneType because of missing information
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
			# Save Excel document
			wb.save('search_results.xlsx')

# Save Excel document
wb.save('search_results.xlsx')

# Close text file
# out.close()