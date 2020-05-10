import requests, json
import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('daas.xlsx')
worksheet = workbook.add_worksheet()

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

content = requests.get("http://43.224.110.34:8081/chart-of-account/v2/api-docs")
json_content = json.loads(content.content)
prefix = json_content["host"] + json_content["basePath"]
print (len(json_content["paths"]))
for paths in json_content["paths"]:
	for api in json_content["paths"][paths]:

		worksheet.write(row, col, prefix + paths)
		print (prefix + paths + ' METHOD ---> ' + api)
		worksheet.write(row, col + 1, api)

		for controller_name in json_content["paths"][paths][api]["tags"]:

			print (controller_name)
			worksheet.write(row, col + 2, controller_name)

			for schema in json_content["paths"][paths][api]['responses']["200"]["schema"]:
				response_type_ref= None;
				if 'items' in schema:
					if api == 'get':
						model_name_raw = json_content["paths"][paths][api]['responses']["200"]["schema"]["items"]["$ref"]
						response_type_main = json_content["paths"][paths][api]['responses']["200"]["schema"]["type"]
						model_name = model_name_raw.split("/")[2]
						model_properties = json_content["definitions"][model_name]["properties"]
						print (model_properties)
						for prop in model_properties:
							if '$ref' in json_content["definitions"][model_name]["properties"][prop]:
								ref_model_raw = json_content["definitions"][model_name]["properties"][prop]["$ref"]
								ref_model_name = ref_model_raw.split("/")[2]

								if 'type' in json_content["definitions"][model_name]["properties"][prop]:
									response_type_ref = json_content["definitions"][model_name]["properties"][prop]["type"]

								if ref_model_name == model_name:
									json_content["definitions"][model_name]["properties"][prop] = '{}'
								else:
									if response_type_ref == 'array':
										json_content["definitions"][model_name]["properties"][prop] = \
											[json_content["definitions"][ref_model_name]["properties"]]
									else:
										json_content["definitions"][model_name]["properties"][prop] = \
											json_content["definitions"][ref_model_name]["properties"]

						res = json_content["definitions"][model_name]["properties"]
						if response_type_main == 'array':
							res = [res]
						worksheet.write(row, col + 3, json.dumps(res, ensure_ascii=False))
						row = row + 1



workbook.close()
