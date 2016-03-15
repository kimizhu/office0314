
# JavaScript API for Office
The JavaScript API for Office includes objects, methods, properties, events, and enumerations that you can use in your Office Add-ins code.

Learn more about [supported hosts and other requirements](../docs/overview/requirements-for-running-office-add-ins.md).

The  **Microsoft.Office.WebExtension** namespace (which by default is referenced using the alias[Office](http://msdn.microsoft.com/library/c490b13d-ee52-4291-af5d-f4a5a11d3af0%28Office.15%29.aspx) in code) contains objects you can use to write script that interacts with content in Office documents, worksheets, presentations, mail items, and projects from your Office Add-ins.
## JavaScript API for Office objects


|**Object**|**Supported add-in type**|**Supported host applications**|
|:-----|:-----|:-----|
|[AsyncResult](http://msdn.microsoft.com/library/540c114f-0398-425c-baf3-7363f2f6bc47%28Office.15%29.aspx)|Content add-in, Outlook add-in, Task pane add-in|Access, Excel, Outlook, PowerPoint, Project, Word|
|[AttachmentDetails](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[Binding](http://msdn.microsoft.com/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e%28Office.15%29.aspx)|Content add-in, Task pane add-in|Access, Excel, Word|
|[Bindings](http://msdn.microsoft.com/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx)|Content add-in, Task pane add-in|Access, Excel, Word|
|[Body](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[Contact](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[Context](http://msdn.microsoft.com/library/662883d5-b86f-4bdc-99f0-9ee9129ed16c%28Office.15%29.aspx)|Content add-in, Outlook add-in, Task pane add-in |Access, Excel, Outlook, PowerPoint, Project, Word|
|[CustomProperties](http://dev.outlook.com/reference/add-ins/CustomProperties.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[CustomXmlNode](http://msdn.microsoft.com/library/dc1518de-47fa-4108-aab7-04a022724b04%28Office.15%29.aspx)|Task pane add-in|Word|
|[CustomXmlPart](http://msdn.microsoft.com/library/83f0e668-8236-4f2f-a20f-b173a9e3f65f%28Office.15%29.aspx)|Task pane add-in |Word|
|[CustomXmlParts](http://msdn.microsoft.com/library/ba40cd4c-29bb-4f31-875d-6f1382fd1ee8%28Office.15%29.aspx)|Task pane add-in |Word|
|[CustomXmlPrefixMappings](http://msdn.microsoft.com/library/18b9aa8c-83e7-4c2f-8530-6a0ac8ce5535%28Office.15%29.aspx)|Task pane add-in |Word|
|[Diagnostics](http://msdn.microsoft.com/library/8ad6a159-ed07-4b82-8897-a80fd208551b%28Office.15%29.aspx)|Outlook add-in|Outlook|
|[Document](http://msdn.microsoft.com/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx)|Content add-in, Task pane add-in|Access, Excel, PowerPoint, Project, Word|
|[EmailAddressDetails](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[EmailUser](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[Entities](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[Error](http://msdn.microsoft.com/library/36d1d048-b888-4bb5-9321-d340bcbc86f4%28Office.15%29.aspx)|Content add-in, Outlook add-in, Task pane add-in|Access, Excel, Outlook, PowerPoint, Project, Word|
|[File](http://msdn.microsoft.com/library/04923ddf-8efa-459f-aed5-d8c06385ca50%28Office.15%29.aspx)|Task pane add-in|PowerPoint, Word|
|[Item](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[Location](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[Mailbox](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[MeetingSuggestion](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[MatrixBinding](http://msdn.microsoft.com/library/35e8568e-9129-4c00-b30f-d8c3b2555f1e%28Office.15%29.aspx)|Content add-in, Task pane add-in|Excel, Word|
|[MeetingSuggestion](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[Office](http://msdn.microsoft.com/library/c490b13d-ee52-4291-af5d-f4a5a11d3af0%28Office.15%29.aspx)|Content add-in, Outlook add-in, Task pane add-in|Access, Excel, Outlook, PowerPoint, Project, Word|
|[PhoneNumber](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[ProjectDocument](http://msdn.microsoft.com/library/1908af4f-93b9-4859-87e3-06942014fae1%28Office.15%29.aspx)|Task pane add-in |Project|
|[Recipients](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[RoamingSettings](http://dev.outlook.com/reference/add-ins/RoamingSettings.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[Settings](http://msdn.microsoft.com/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx)|Content add-in, Task pane add-in|Access, Excel, PowerPoint, Word|
|[Slice](http://msdn.microsoft.com/library/011b5647-639b-4b06-8625-ba9de01bed4b%28Office.15%29.aspx)|Task pane add-in|PowerPoint, Word, Word Online|
|[Subject](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[TableBinding](http://msdn.microsoft.com/library/1508795b-1c70-456c-b3bf-666d40cf8f50%28Office.15%29.aspx)|Content add-in, Task pane add-in|Access, Excel, Word|
|[TableData](http://msdn.microsoft.com/library/2183ea52-5a40-4048-b9a4-7cd66bb0ad5d%28Office.15%29.aspx)|Content add-in, Task pane add-in|Access, Excel, Word|
|[TaskSuggestion](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[TextBinding](http://msdn.microsoft.com/library/6b71b21d-f64d-425c-99d9-c62b2a9969be%28Office.15%29.aspx)|Content add-in, Task pane add-in|Excel, Word|
|[Time](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)|Outlook add-in|Outlook|
|[UserProfile](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.userProfile.html%28Office.15%29.md)|Outlook add-in|Outlook|

## Enumerations

|**Parent topic**|**Supported add-in type**|**Supported host applications**|
|:-----|:-----|:-----|
|[Enumerations](http://msdn.microsoft.com/library/eee5e332-6d83-4b58-974d-3abe002f4359%28Office.15%29.aspx)|See child enumeration topics for details.|See Requirements in enumeration topic for details.|

## View APIs by add-in type support

To view the JavaScript API for Office organized by the subsets of the API that support each add-in type, see

|**API **|**Description**|
|:-----|:-----|
|[Shared API](http://msdn.microsoft.com/library/4f21922c-bf0d-4617-9071-9c99413f4977%28Office.15%29.aspx)|The subset of the API that you can use in all three types of Office Add-ins: content, task pane, and Outlook add-ins.|
|[Document API](http://msdn.microsoft.com/library/1bad7bff-2161-46c6-b536-eb4a0608b7ac%28Office.15%29.aspx)|The subset of the API that you can use in the two types of Office Add-ins associated with documents: content and task pane add-ins.|
|[Mailbox API](http://dev.outlook.com/reference/add-ins/index.mdl.aspx)|The subset of the API that you can use in Outlook add-ins.|

## Supported host applications
* Access
* Excel
* Outlook
* PowerPoint
* Project
* Word
