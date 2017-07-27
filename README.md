# SharePoint Embed Shortcodes
Embed publicly shared documents from SharePoint and OneDrive for Business in WordPress

Currently shortcodes are only provided to embed Excel documents; Word and PowerPoint documents should be supported,
but are waiting on Microsoft to provide documentation.

## Excel Shortcode Documentation
The Excel embed shortcode is `[sp-excel-embed]`. It accepts the following attributes:

* `source` : accepts public sharing URL (default: false)
* `width` : pixel or precentage value (default: 100%)
* `height` : pixel or precentage value (default: 500px)
* `bifeatures` : enable BI features such as Power View visualizations, PivotTables, and Data Model-based slicers work in the embedded workbook (default: true)
* `interactive` : enable interactivity with filters and Pivot tables in the workbook (default:false)
* `download` : show integrated download link

Example shortocde with default values:
```
[sp-excel-embed source="https://bellevuec.sharepoint.com/sites/its/web/_layouts/15/guestaccess.aspx?docid=XXXXX&authkey=XXXXXX" width="100%" height="500px" bifeatures="true" interactive="false" download="true"]
```

### How to get source URL
SharePoint and OneDrive for Business allow you to create sharing links in a variety of ways; any of these will work as long as the settings are correct. 

Primary method:
1. Start from the Excel Online 'read only' view
2. Click **Share** then **Share with people**
3. Click **Get a link** 
4. Select **View Link - no sign-in required**. Warning: Make sure that the link does **NOT** allow for edit access!
5. Select and copy the link provided to the clipboard
6. Click **Close**

## Microsoft Documentation
See this [Office Support Article](https://support.office.com/en-us/article/Embed-your-Excel-workbook-on-your-web-page-or-blog-from-SharePoint-or-OneDrive-for-Business-7af74ce6-e8a0-48ac-ba3b-a1dd627b7773) for documentation on how to configure an Excel embed. 
