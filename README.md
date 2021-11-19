# FixListForm 

When user accidentally removed the SharePoint List form from any type of Lists. It is difficult to revert the changes. This script works on classic UX only.
This script is to fix those forms by adding additional webpart or recreate the form pages.

Those are the parameters.

-siteUrl … Site (Web) URL

-listName … List Title

-username … (optional) Site Collection Admin user who fixes the page.

-password … (optional) The password of the above user.

-formtype … (optional) The target form type. DISPLAY, NEW, EDIT, ALL (The default option is ALL)

-force … (optional) Recreate the page when $true is specified. Otherwise, adding list form webpart on that page.

## Reference
Original Code is from Japan SharePoint Support Team Blog.
URL: https://blogs.technet.microsoft.com/sharepoint_support/2015/11/27/sharepoint-online-125/
