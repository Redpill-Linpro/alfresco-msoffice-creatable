/**
* Redpill Linpro root namespace.
* 
* @namespace Extras
*/
if (typeof RPLP == "undefined" || !RPLP)
{
   var RPLP = {};
}


/**
 * LibreOffice client script
 * 
 * @namespace RPLP
 * @class RPLP.MsOfficeCreateNewDocument
 */
(function()
{	

   /**
    * YUI Library aliases
    */
   var Dom = YAHOO.util.Dom;
   
   /**
    * Alfresco Slingshot aliases
    */
   var $html = Alfresco.util.encodeHTML,
      $links = Alfresco.util.activateLinks,
      $combine = Alfresco.util.combinePaths,
      $userProfile = Alfresco.util.userProfileLink,
      $siteURL = Alfresco.util.siteURL,
      $date = function $date(date, format) { return Alfresco.util.formatDate(Alfresco.util.fromISO8601(date), format); },
      $relTime = Alfresco.util.relativeTime,
      $isValueSet = Alfresco.util.isValueSet;
      
   /**
    * DocumentGeographicInfo constructor.
    * 
    * @param {String} htmlId The HTML id of the parent element
    * @return {Extras.DocumentGeographicInfo} The new DocumentGeographicInfo instance
    * @constructor
    */
   RPLP.MsOfficeCreateNewDocument = function(htmlId)
   {
	   RPLP.MsOfficeCreateNewDocument.superclass.constructor.call(this, "RPLP.MsOfficeCreateNewDocument", htmlId);
      
      return this;
   };

	 /**
	* Augment prototype with Actions module
	*/
   /**
    * Augment prototype with Actions module
    */
   YAHOO.lang.augmentProto(RPLP.MsOfficeCreateNewDocument, Alfresco.doclib.Actions);
   
   YAHOO.extend(RPLP.MsOfficeCreateNewDocument, Alfresco.component.Base,
   {
       /**
        * Object container for initialization options
        *
        * @property options
        * @type object
        */
       options:
       {
           /**
            * Reference to the current site
            * 
            * @property siteId
            * @type string
            * @default null
            */
           siteId: null,
           
           createContentActions: []

      },
      
      docListToolbar: null,
      
      /**
       * Event handler called when "onReady"
       *
       * @method: onReady
       */
      onReady: function MsOfficeCreateNewDocument_onReady(layer, args)
      {
    	  var parent = this;
    	  
    	  
    	  //'this.docListToolbar.doclistMetadata.onlineEditing' contains true or false if online editing is on
    	  
    	  //This registers an action for each menu item, not one common
    	  // actionName is different, and this is also registered in toolbar.get.js
    	  // this was the one way to get custom data through Alfresco code without rewriting, keeping 4.2.x style extension
    	  for (var j = 0; j < this.options.createContentActions.length; j++)
          {
    		  (function(actionDefinition) { // get a local copy of the current value
    			  YAHOO.Bubbling.fire("registerAction",
    			    {
		    		  actionName: "onNewOfficeDocument" + actionDefinition.index,
		    	        fn: function rplp_onNewOfficeDocument(file) {
		    	        	parent.docListToolbar = this;
		    	            if (parent._launchOnlineCreator(file, actionDefinition))
		    	            {
		    	               YAHOO.Bubbling.fire("metadataRefresh");
		    	            }
		    	            else
		    	            {
		    	               Alfresco.util.PopupManager.displayMessage(
		    	               {
		    	                  text: this.msg("message.edit-online.office.failure")
		    	               });		 
		    	            }
		    	        }
    			    });
          		})(this.options.createContentActions[j]);
          }
 
      },
  	 
      /**
       * Opens the appropriate Microsoft Office application for online editing.
       * Supports: Microsoft Office 2003, 2007 & 2010.
       *
       * @method Alfresco.util.sharePointOpenDocument
       * @param record {object} Object literal representing file or folder to be actioned
       * @return {boolean} True if the action was completed successfully, false otherwise.
       */
      _copyNode: function dlA__copy(newParentNodeRef, sourcePath, loc)
      {

    	  var me = this;
           Alfresco.util.Ajax.jsonPost(
                   {
                      url: Alfresco.constants.PROXY_URI + "msoffice-creatable/copyin",
                      dataObj:
                      {
                          sourcePath: sourcePath,
                          parentNodeRef: newParentNodeRef
                      },
                      successCallback:
                      {
                    	  
                         fn: function (response)
                         {
                        	 

                            
                            if (null != response) {
                            	//fetch resulting node
                          	   var newName = response.json.name;
                  	    	   // new name
                          	   var docList = Alfresco.util.ComponentManager.findFirst("Alfresco.DocumentList");
                          	   var docToolbar = this.docListToolbar;
                          	   
                          	   
                          	    //YAHOO.Bubbling.subscribe("highlightFile", function RPLP_onFilterChanged(layer, args){
                          	   docList.widgets.dataTable.subscribe("renderEvent", function RPLP_onFilterChanged(layer, args){ 
                          		// edit online
                            	   if (docToolbar.doclistMetadata.onlineEditing) {
                            		 var editOnlineUrl = Alfresco.util.onlineEditUrl(docToolbar.doclistMetadata.custom.vtiServer, loc);
                            		 // Reference to Document List component
                                  
                            		  //var record = this.widgets.dataTable.getRecord(elIdentifier).getData(); 
                            	     var records = docList.widgets.dataTable.getRecordSet().getRecords();
                            	     for (var ri = 0; ri < records.length; ri++)
                            	     { 
                            	    	rData = records[ri].getData();
                            	    	if (rData.node.type == "cm:content" && rData.fileName == newName) {
                            	    		docToolbar._launchOnlineEditor(rData);
                            	    		docList.widgets.dataTable.unsubscribe("renderEvent", arguments.callee);
                            	    		break;
                            	    	}
                            	      }
                            	    }
                            	   
                          	  // Make sure we get other components to update themselves to show the new content
                                YAHOO.Bubbling.fire("nodeCreated",
                                {
                                   name: response.json.name,
                                   parentNodeRef: newParentNodeRef,
                                   highlightFile: response.json.name
                                });
                                
                         	   
                         	
//                           	var foo = {
//                           		  count :0,
//                           		  'method' : function(data) {
//                           		    this.count++;
//                           		    if(this.count == 10) {
//                           		      timer.cancel();
//                           		    }
//                           		    console.log(this.count);
//                           		  }
//                           		}
//                           		var timer = YAHOO.lang.later(1000, foo, 'method', [{data:'bar', data2:'zeta'}], true);`
                         			 
//                         	YAHOO.Bubbling.on("filterChanged",
//                         			YAHOO.Bubbling.unsubscribe("filterChanged", this.onBeforeFormRuntimeInit); 
                            }, this);
                         	   
                 	    	   
                            } 
                         },
                         scope: this,
                         loc: loc
                      	 
                      },
                      successMessage: this.msg("message.create-content-by-template-node.success", sourcePath),
                      failureMessage: this.msg("message.create-content-by-template-node.failure", sourcePath)
                   });
		 
      },
    /**
     * Opens the appropriate Microsoft Office application for online editing.
     * Supports: Microsoft Office 2003, 2007 & 2010.
     *
     * @method Alfresco.util.sharePointOpenDocument
     * @param record {object} Object literal representing file or folder to be actioned
     * @return {boolean} True if the action was completed successfully, false otherwise.
     */
    _launchOnlineCreator: function dlA__launchOnlineCreator(record, actionDefinition)
    {
       var controlProgID = "SharePoint.OpenDocuments", //expression.CreateNewDocument2(pdisp, bstrTemplateLocation, bstrDefaultSaveLocation)
          jsNode = record.jsNode,
          loc = {
    		   path: this.docListToolbar.currentPath,
    		   site: {
    			   name: this.options.siteId
    		   },
    		   container: {
    			   name: this.docListToolbar.options.containerId
    		   },
    		   file: ((null != actionDefinition.filename) ? actionDefinition.filename : "newDocument." + actionDefinition.icon) //what to do, it does not matter, Office will not succeed with this and we get a directory save box!
       	},
          appProgID = actionDefinition.msTemplate,
          activeXControl = null;
       if (!this.docListToolbar.doclistMetadata.onlineEditing || "true" == actionDefinition.copyin) {

    	   var newParentNodeRef = record.nodeRef; 
    	   //copy & launch edit of doc
           this._copyNode(newParentNodeRef, appProgID, loc);
           return true;
       } 
       else if (appProgID !== null)
       {
    	   //Check if a path is given (for a template)
    	   // if path MS Office will retrieve the template as basis for new document
    	   if (-1 < appProgID.indexOf("/")) {
    		   var vtiServer = this.docListToolbar.doclistMetadata.custom.vtiServer;
    		   appProgID = vtiServer.protocol + "://" + vtiServer.host + ":" + vtiServer.port + "/" + vtiServer.contextPath 
    		   + ((appProgID.substring(0,1) == "/")? "" : "/")
    		   + appProgID;
    		   
    	   }
          // Ensure we have the record's onlineEditUrl populated
          if (!$isValueSet(record.onlineEditUrl))
          {
             record.onlineEditUrl = Alfresco.util.onlineEditUrl(this.docListToolbar.doclistMetadata.custom.vtiServer, loc);
          }

          if (YAHOO.env.ua.ie > 0)
          {
             return this._launchOnlineCreatorIE(controlProgID, record, appProgID);
          }
          
          if (Alfresco.util.isSharePointPluginInstalled())
          {
             return this._launchOnlineCreatorPlugin(record, appProgID);
          }
          else
          {
             Alfresco.util.PopupManager.displayPrompt(
             {
                text: this.msg("actions.editOnline.failure", loc.file)
             });
             return false;
          }
       }

       // No success in launching application via ActiveX control; launch the WebDAV URL anyway
       return window.open(record.onlineEditUrl, "_blank");
    },

    /**
     * Opens the appropriate Microsoft Office application for online editing.
     * Supports: Microsoft Office 2003, 2007 & 2010.
     *
     * @method Alfresco.util.sharePointOpenDocument
     * @param record {object} Object literal representing file or folder to be actioned
     * @return {boolean} True if the action was completed successfully, false otherwise.
     */
    _launchOnlineCreatorIE: function dlA__launchOnlineCreatorIE(controlProgID, record, appProgID)
    {
       // Try each version of the SharePoint control in turn, newest first
       try
       {
//          if (appProgID === "Visio.Drawing")
//             throw ("Visio should be invoked using activeXControl.EditDocument2.");
          activeXControl = new ActiveXObject(controlProgID + ".3");
          return activeXControl.CreateNewDocument2(window, appProgID, record.onlineEditUrl);
       }
       catch(e)
       {
          try
          {
             activeXControl = new ActiveXObject(controlProgID + ".2");
             return activeXControl.CreateNewDocument(appProgID, record.onlineEditUrl);
          }
          catch(e1)
          {
             try
             {
                activeXControl = new ActiveXObject(controlProgID + ".1");
                return activeXControl.CreateNewDocument(appProgID, record.onlineEditUrl);
             }
             catch(e2)
             {
                // Do nothing
             }
          }
       }
       return false;
    },


    
    /**
     * Opens the appropriate Microsoft Office application for online editing.
     * Supports: Microsoft Office 2010 & 2011 for Mac.
     *
     * @method Alfresco.util.sharePointOpenDocument
     * @param record {object} Object literal representing file or folder to be actioned
     * @return {boolean} True if the action was completed successfully, false otherwise.
     */
    _launchOnlineCreatorPlugin: function dlA__launchOnlineCreatorPlugin(record, appProgID)
    {
       var plugin = document.getElementById("SharePointPlugin");
       if (plugin == null && Alfresco.util.isSharePointPluginInstalled())
       {
          var pluginMimeType = null;
          if (YAHOO.env.ua.webkit && Alfresco.util.isBrowserPluginInstalled("application/x-sharepoint-webkit"))
             pluginMimeType = "application/x-sharepoint-webkit";
          else
             pluginMimeType = "application/x-sharepoint";
          var pluginNode = document.createElement("object");
          pluginNode.id = "SharePointPlugin";
          pluginNode.type = pluginMimeType;
          pluginNode.width = 0;
          pluginNode.height = 0;
          pluginNode.style.setProperty("visibility", "hidden", "");
          document.body.appendChild(pluginNode);
          plugin = document.getElementById("SharePointPlugin");
          
          if (!plugin)
          {
             return false;
          }
       }
       
       try
       {
//          if (appProgID === "Visio.Drawing") 
//             throw ("Visio should be invoked using activeXControl.NewDocument2.");
          return plugin.CreateNewDocument2(window, appProgID, record.onlineEditUrl);
       }
       catch(e)
       {
          try
          {
             return plugin.CreateNewDocument(appProgID, record.onlineEditUrl);
          }
          catch(e1)
          {
                return false;
          }
       }
    }
    });
  
})();
  



      