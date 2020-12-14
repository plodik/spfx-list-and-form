## Sharepoint Framework: Complex LIST and FORM sample

The following sample List and Form Sharepoint Framework (SPFx) component is based on SPFx 1.11.0 for Sharepoint Online using React.

It demonstrates the complex List and Form scenario for Sharepoint List items. We frequently require to provide a web part with advanced List with filtering, paging and custom column selection, similar to native Sharepoint list experience. The second part is Form with advanced controls like DatePicker, PeoplePicker, custom external REST calls, validation, PDF generation for printing/archiving purposes, etc.

It is a fully functioning and working component, ready to be customized and deployed. Most of the other online samples, documents and tutorials provide just incomplete parts of the whole solution. I would like to provide you with a full package when all is connected together.

The solution contains 2 main SPFx components, optimized in Responsive design for Modern SharePoint Sites:

### LIST
Configurable responsive list of SPListItems:

	* Selectable columns from the source SPList with configurable order
	* Filtering
	* Sorting by column names
	* Paging

![List web part with config](https://github.com/plodik/spfx-list-and-form/blob/master/!Screenshots/List%20with%20config.png)
	
### FORM
Complex form component for Add or Edit action:

* Includes custom PeoplePicker
* Includes custom DatePicker
* Includes dropdown for Lookup column
* Advanced validation of required fields
* Sample call to external REST API resource
* Approval process based on State lookup list
* When item is "approved" (last state is set), custom PDF can be generated in browser for printing or archive purposes
* When the page is called without querystring, New Item Form is used. When call with ?A=Edit&poid=XXX, where XXX is ID of existing list item, the Edit Form is used.

![New item form](https://github.com/plodik/spfx-list-and-form/blob/master/!Screenshots/Form%20Add%20New%20blank.png)

![Edit item form with Ask for approval step](https://github.com/plodik/spfx-list-and-form/blob/master/!Screenshots/Form%20Edit%20with%20Ask%20For%20Approval.png)
	
### External packages in the solution:

*  PnP libraries for SPO Lists: https://pnp.github.io/pnpjs/
*  Fluent UI React: https://www.npmjs.com/package/office-ui-fabric-react
*  React JS Pagination: https://www.npmjs.com/package/react-js-pagination
*  jsPDF: https://github.com/MrRio/jsPDF

### Setup instructions
App Catalog: Ensure the App Catalog is setup in your SharePoint Online tenant (https://docs.microsoft.com/en-us/sharepoint/use-app-catalog)

The solution is not perfect in terms of code quality, architecture or performance. Further tweaking and improvements in those areas are needed. Most of the sections should be optimized. However it is a working and complex sample of custom list and custom form web part.

### Disclaimer:
The solution should be used only for educational purposes. There is no warranty for the production usage as is. 
We grant you a nonexclusive, royalty-free right to use and modify the Sample Code as needed.