# Overview
A quick and dirty PowerShell script, which reads document metadata from a CSV, matches each entry to the corresponding PDF file, and import the documents - including all metadata - via the Paperless-ngx API.
Would be created for Paperless-ngx version 2.14.7

# Custom fields
#Custom fields in Paperless-ngx:
The script assumes, that in Paperless-ngx are custom fields with the following IDs:

| Custom field ID | Custom field name |
|-----------------|------------------|
| 1               | Company          |
| 2               | Amount           |
| 3               | InvoiceDate      |
| 4               | InvoiceID        |
| 5               | Notes            |