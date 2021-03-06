
# Method element
Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.

 **Add-in type:** Content, Task pane


## Syntax:


```XML
<Method Name="string "/>
```


## Contained in:

 _ [Methods](../../reference/manifest/methods.md)_


## Attributes



|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|Name|string|required|Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.|

## Remarks

The  **Methods** and **Method** elements aren't supported by mail add-ins. For more information about requirement sets, see [Specify Office hosts and API requirements](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx#SpecifyRequirementSets_intro).


 >**Important**  Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an  **if** statement when calling that method in the script of your add-in. For more information about how to do this, see [Understanding the JavaScript API for Office](http://msdn.microsoft.com/library/01180dae-ca45-40c8-b3dd-fd2a85651c0c%28Office.15%29.aspx#HostAPISupport_UsingIfStatements).

