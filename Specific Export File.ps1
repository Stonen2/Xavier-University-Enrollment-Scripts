#SPECIFIC EXPORT FILE DETAILS
##GET /api/v1/exports/:export_id/export_files/:export_file_id

#User name to sign into the webadmit
$user_id = "" #get one from User Identity endpoint

#Name of the data we are trying to pull
$export_id = "" #get one from Export endpoint


#The name of the file we are trying to pull from webadmits
$export_file_id = "" #get one from Export Files endpoint


#This is the API call we want to use to get the json file
$api_key = ""
$header = @{"x-api-key"=$api_key}

#The url of the webadmit report that we are trying to accesss
$uri = "https://api.webadmit.org/api/v1/exports/$export_id/export_files/$export_file_id"

#
#This part of the code invokes the web request to import the excel file 
#From this file we get a JSON file from webadmit. 
#We then convert the JSON into xls
#THen we download the excel file 


$resp = invoke-webrequest -uri $uri -method GET -headers $header
$obj = convertfrom-json -inputobject $resp.content
$obj.export_files | format-table
$download_uri = $obj.export_files.download_url
