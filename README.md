SharePoint Hauler is a desktop application for SharePoint power-users. The application allows you to copy the contents of SharePoint Libraries and Lists between SharePoint Web Sites, Site Collections or Tenants while preserving the Metadata associated with the content and the Folder Structure of the content.

####SUPPORTED VERSIONS:

SharePoint 2010

SharePoint 2013

SharePoint 2016

SharePoint Online

Copying of the content of Lists/Libraries is possible between any of the above SharePoint Versions.
Also content from any of the above SharePoint Versions can be backed up to a File System and restored later.

####GENERAL SCENARIOS

•	Copy Content (List Items and Library files) fractionally or as a whole between any version of SharePoint while preserving the Metadata associated with the content and the Folder Structure of the content.

•	Copy Content (Library Files) to File System while preserving the Metadata associated with the content and the Folder Structure of the content.

•	Upload files from File System to SharePoint while preserving the Folder Structure of the content.

•	Restore previously copied Content with the Application from File System to SharePoint while preserving the Metadata associated with the content and the Folder Structure of the content.

####METADATA CONSIDERATIONS

The following table shows how the Application handles all types of SharePoint fields and if there are any considerations before copying.

Single line of text:			No considerations

Multiple lines of text:			No considerations

Choice:							The field should be configured the same as Source on Destination

Number:							No considerations

Currency:						The field should be configured the same as Source on Destination

Date and Time:					No considerations

Lookup:							No considerations if copying in the same Site Collection. Can be copied as Text in other scenarios

Yes/No (check box):				No considerations

Person or Group:				No considerations if copying in the same Active Directory. Can be copied if User Mapping is provided in other scenarios

Hyperlink or Picture:			No considerations

Calculated:						Will be automatically filled when dependent fields are copied. If these dependent fields are “Lookup” or “Metadata” fields reconfiguration of the Calculated field should be done.

Task Outcome:					The field should be configured the same as Source on Destination

External Data:					Not supported

Managed Metadata:				No considerations if copying in the same Farm/Tenant. Can be copied as Text in other scenarios

####CROSS DOMAIN COPY OF CONTENT AND USER FIELDS

The Application provides a way to preserve “People and Groups” fields when copying content across Domains/Tenants. For such a scenario a CSV file should be prepared which maps the Source Users to the Destination Users. The CSV should be in the following format – two columns, one named Source_User and the other Destination_User. 

####VERSIONING

The Application copies only the latest version of each Item/Document. 

####LIST ITEM ATTACHMENTS

The Application copies all List Item Attachments automatically.

