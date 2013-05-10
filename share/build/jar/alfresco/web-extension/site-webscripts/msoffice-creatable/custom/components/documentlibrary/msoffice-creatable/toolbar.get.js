function sortByIndex(obj1, obj2)
{
   return (obj1.index > obj2.index) ? 1 : (obj1.index < obj2.index) ? -1 : 0;
}

if (model.widgets)
{
   for (var i = 0; i < model.widgets.length; i++)
   {
      var widget = model.widgets[i];
      if (widget.id == "DocListToolbar")
      { 
    	  var newCreateContent = [];
    	  
    	   var msoConfig = config.scoped["DocumentLibrary"]["msoffice-creatable"];

    	   if (msoConfig !== null)
    	   {

    		   
               var configs = msoConfig.getChildren("creatable-types"),
    	            creatableConfig,
    	            configItem,
    	            creatableType,
    	            fileName,
    	            msTemplate,
    	            copyin,
    	            index,
    	            url;

    	         if (configs)
    	         {
    	            for (var i = 0; i < configs.size(); i++)
    	            {
    	               creatableConfig = configs.get(i).childrenMap["creatable"];
    	               if (creatableConfig)
    	               {
    	            	   
    	                  for (var j = 0; j < creatableConfig.size(); j++)
    	                  {
    	                     configItem = creatableConfig.get(j);
    	                     // Get type and mimetype from each config item
    	                     creatableType = null != configItem.attributes["type"] ? configItem.attributes["type"].toString() : "doc";
    	                     fileName = null != configItem.attributes["filename"] ? configItem.attributes["filename"].toString() : "new";
    	                     copyin = null != configItem.attributes["use-copy-in"] ? configItem.attributes["use-copy-in"].toString() : "false";   
    	                     msTemplate = configItem.value.toString();
    	                     index = parseInt(configItem.attributes["index"] || "0");
    	                     if (creatableType && msTemplate)
    	                     {
    	                       // url = "create-content?destination={nodeRef}&itemId=cm:content&formId=doclib-create-googledoc&mimeType=" + mimetype;
    	                        
    	                        newCreateContent.push(
    	                        {
    	                           type: "javascript",
    	                           icon: creatableType,
    	                           filetype: creatableType,
    	                           filename: fileName, 
    	                           label: "msoffice." + creatableType,
    	                           msTemplate: msTemplate,
    	                           copyin: copyin,
    	                           index: index,
    	                           permission: "CreateChildren",
    	                           params:
    	                           {
    	                              "function": "onNewOfficeDocument" + index
    	                           }
    	                        });
    	                     }
    	                  }
    	               }
    	            }
    	         }
    	   }
    	   
         var oldConditions = widget.options.createContentActions;//eval("(" + widget.options.createContentActions + ")");
         // Add the other conditions back in
         for (var j = 0; j < newCreateContent.length; j++)
         {
        	 oldConditions.push(newCreateContent[j]);
         }
         // Override the original conditions
         model.createContent = oldConditions.sort(sortByIndex);
         widget.options.createContentActions = model.createContent;
      }
   }
   var msOfficeCreatable = {
      id: "MsOfficeCreatable", 
      name: "RPLP.MsOfficeCreateNewDocument",
      options: {
         siteId: (page.url.templateArgs.site != null) ? page.url.templateArgs.site : "",
         createContentActions: newCreateContent.sort(sortByIndex)
      }
   };
   model.widgets.push(msOfficeCreatable);
}