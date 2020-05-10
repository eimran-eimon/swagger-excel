import requests
import xlsxwriter
import json
from pprint import pprint
import jsonref
import xlsxwriter
from jsonpath_ng import jsonpath, parse

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('daas.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column(0, 4, 30)
header_format = workbook.add_format({
	'bold': True,
	'text_wrap': True,
	'align': 'center',
	'fg_color': '#D7E4BC',
	'border': 1})


def make_obj_or_array(body):
	if 'type' in body:
		if body["type"] == 'array':
			body = [body["items"]]
			return body
		elif body["type"] == 'object':
			body = body["properties"]
			return body


# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

content = requests.get("http://43.224.110.34:8081/chart-of-account/v2/api-docs")
json_ref = jsonref.loads(content.content)
json_content = json.loads(content.content)
prefix = json_content["host"] + json_content["basePath"]

type_of_controllers = [match.value for match in parse('$.tags[*].name').find(json_content)]
type_of_controllers.remove('basic-error-controller')
list_of_controller = [match.value for match in parse('$.paths[*]..tags[:1]').find(json_content)]
api_list = [match.value for match in parse('$.paths.*.*').find(json_content)]
api_list_with_ref = [match.value for match in parse('$.paths.*.*').find(json_ref)]
original_path = [str(match.full_path) for match in parse('$.paths.*.*').find(json_content)]
pprint(original_path)

# Add a bold format to use to highlight cells.
center = workbook.add_format({'align': 'center'})
worksheet.write(row, col, 'Controller Name', header_format)
worksheet.write(row, col + 1, 'Method', header_format)
worksheet.write(row, col + 2, 'Task', header_format)
worksheet.write(row, col + 3, 'Url', header_format)
worksheet.write(row, col + 4, 'JSON Data', header_format)

row = row + 1
for type_of_controller in type_of_controllers:
	indices = [i for i, x in enumerate(list_of_controller) if x == type_of_controller]
	print (indices)
	for index in indices:
		try:
			worksheet.write(row, col, type_of_controller, center)
			api_description = api_list[index]
			path = original_path[index]
			worksheet.write(row, col + 1, path.split(".")[2].upper(), center)
			worksheet.write(row, col + 2, api_description["operationId"])
			worksheet.write(row, col + 3, prefix + path.split(".")[1])
			raw_body = api_list_with_ref[index]["responses"]["200"]["schema"]
			body = make_obj_or_array(raw_body)
			print ("-------------------")
			if type(body) == list:
				body = [make_obj_or_array(body[0])]
			worksheet.write(row, col + 4, str(body))
			row = row + 1
		except KeyError:
			print ('Not Found!')
	row = row + 1


workbook.close()

# pprint(json_content["paths"]["/economic/code"]["post"]["parameters"])
