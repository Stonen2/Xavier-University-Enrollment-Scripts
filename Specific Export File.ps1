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






############################################################################
#USER IDENTITY

#Use the proper API Call to then access the URL within WebAdmit that holds the data
$api_key = ""
$header = @{"x-api-key"=$api_key}
$uri = "https://api.webadmit.org/api/v1/user_identities"


#
#This part of the code invokes the web request to import the excel file 
#From this file we get a JSON file from webadmit. 
#We then convert the JSON into xls
#THen we download the excel file 
$resp = invoke-webrequest -uri $uri -method get -headers $header

$obj = convertfrom-json -inputobject $resp.content

$obj.user_identities | format-table -property id,cycle,association






#####################################################################################

#LIST OF EXPORTS
##GET /api/v1/user_identities/:user_identity_id/exports

#User ID to log in
$user_id = "" #get one from User Identity endpoint

#Use the proper API Call to then access the URL within WebAdmit that holds the data
$api_key = ""
$header = @{"x-api-key"=$api_key}
$uri = "https://api.webadmit.org/api/v1/user_identities/$user_id/exports"


#
#This part of the code invokes the web request to import the excel file 
#From this file we get a JSON file from webadmit. 
#We then convert the JSON into xls
#THen we download the excel file 
$resp = invoke-webrequest -uri $uri -method get -headers $header
$obj = convertfrom-json -inputobject $resp.content
$obj.exports | format-table -property id,name,list_type,format,mime_type




######################################################################################



#LIST OF PDF MANAGER TEMPLATES
$user_id = "" #get one from User Identity endpoint
$api_key = ""
$header = @{"x-api-key"=$api_key}
$uri = "https://api.webadmit.org/api/v1/user_identities/$user_id/pdf_manager_templates"
$resp = invoke-webrequest -uri $uri -method GET -headers $header
$obj = convertfrom-json -inputobject $resp.content
$obj.pdf_manager_templates | format-table -property id,name,list_name,document_source_title,href


#################################################################################

#LIST OF AVAILABLE PDF BATCHES
##list all existing batches GET /api/v1/user_identities/:user_identity_id/pdf_manager_batches
$user_id = "" #get one from User Identity endpoint
$api_key = ""
$header = @{"x-api-key"=$api_key}
$uri = "https://api.webadmit.org/api/v1/user_identities/$user_id/pdf_manager_batches"
$resp = invoke-webrequest -uri $uri -method GET -headers $header
$obj = convertfrom-json -inputobject $resp.content
$batch = $obj.pdf_manager_batches
$template = $batch.pdf_manager_template
$batch | format-table
$template | format-table


##################################################################################
#SPECIFIC PDF BATCH DETAILS
$user_id = "" #get one from User Identity endpoint
$batch_id = "" #get one from Batch endpoint
$api_key = ""
$header = @{"x-api-key"=$api_key}
$uri = "https://api.webadmit.org/api/v1/user_identities/$user_id/pdf_manager_batches/$batch_id"
$resp = invoke-webrequest -uri $uri -method GET -headers $header
$obj = convertfrom-json -inputobject $resp.content
$batch = $obj.pdf_manager_batch
$template = $batch.pdf_manager_template
$download = $batch.download_hrefs
$batch | format-table
$template | format-table
$download | format-table

