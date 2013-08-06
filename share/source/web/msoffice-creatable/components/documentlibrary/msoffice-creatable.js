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
    	  //this ~ documentLibrary
    	  var parent = this;
    	  
    	  
    	  //'this.docListToolbar.doclistMetadata.onlineEditing' contains true or false if online editing is on
    	  
    	  //This registers an action for each menu item, not one common
    	  // actionName is different, and this is also registered in toolbar.get.js
    	  // this was the one way to get custom data through Alfresco code without rewriting, keeping 4.2.x style extension
    	  for (var j = 0; j < this.options.createContentActions.length; j++)
          {
    		  // if ("false" == this.options.createContentActions[j].submenu) {
	    		  (function(actionDefinition) { // get a local copy of the current value
	    			  YAHOO.Bubbling.fire("registerAction",
	    			    {
			    		    actionName: "onNewOfficeDocument" + actionDefinition.index,
			    	        fn: function rplp_onNewOfficeDocument(file)
	    			  		{
			    	        	//this ~ DocumentList toolbar
			    	          	parent.docListToolbar = this;
			    	          	parent.newOfficeDocument(file, actionDefinition);
			    	        }
			    	          	
	    			    });
	          		})(this.options.createContentActions[j]);
    		  //}
          }
 
    	  //monkey path
    	  /**
    	   * appends to submenu
    	   * 
    	   */ 
    	  (function(origBeforeShow, createContentActions) {
    		  Alfresco.DocListToolbar.prototype.onCreateByTemplateNodeBeforeShow = function() {
    			  
    			  // Display loading message
    		         var templateNodesMenu = this.widgets.createContent.getMenu().getSubmenus()[0];
    		         if (templateNodesMenu.getItems().length == 0)
    		         {
    		            templateNodesMenu.clearContent();
    		            templateNodesMenu.addItem(this.msg("label.loading"));
    		            templateNodesMenu.render();

    		            // Load template nodes
    		            Alfresco.util.Ajax.jsonGet(
    		            {
    		               url: Alfresco.constants.PROXY_URI + "slingshot/doclib/node-templates",
    		               successCallback:
    		               {
    		                  fn: function(response, menu)
    		                  {
    		                     var nodes = response.json.data,
    		                        menuItems = [],
    		                        name;
    		                     for (var i = 0, il = nodes.length; i < il; i++)
    		                     {
    		                        node = nodes[i];
    		                        name = $html(node.name);
    		                        if (node.title && node.title !== node.name && this.options.useTitle)
    		                        {
    		                           name += '<span class="title">(' + $html(node.title) + ')</span>';
    		                        }
    		                        menuItems.push(
    		                        {
    		                           text: '<span title="' + $html(node.description) + '">' + name +'</span>',
    		                           value: node
    		                        });
    		                     }
    		                  // Make sure we load sub menu lazily with data on each click
    		                     var createContentMenu = this.widgets.createContent.getMenu(),
    		                         groupIndex = 0;
    		                     //Added
    		                     for (var j = 0; j < createContentActions.length; j++)
    		                     {
	    		               		  if ("true" == createContentActions[j].submenu) {
	    		                          var menuItem, content, url, config, html, li;
	    		                          
	    		                             // Create menu item from config
	    		                             content = createContentActions[j];
	    		                             config = { parent: createContentMenu };
	    		                             url = null;

    		                                config.onclick =
    		                                {
    		                                   fn: function(eventName, eventArgs, obj)
    		                                   {
    		                                      // Copy node so we can safely pass it to an action
    		                                      var node = Alfresco.util.deepCopy(this.doclistMetadata.parent);

    		                                      // Make it more similar to a usual doclib action callback object
    		                                      var currentFolderItem = {
    		                                         nodeRef: node.nodeRef,
    		                                         node: node,
    		                                         jsNode: new Alfresco.util.Node(node)
    		                                      };
    		                                      this[obj.params["function"]].call(this, currentFolderItem);
    		                                   },
    		                                   obj: content,
    		                                   scope: this
    		                                };

    		                                url = '#';
	    		             
	    		                             // Create menu item
	    		                             html = '<a href="' + url + '" rel="' + content.permission + '"><span style="background-image:url(' + Alfresco.constants.URL_RESCONTEXT + 'components/images/filetypes/' + content.icon + '-file-16.png)" class="' + content.icon + '-file">' + this.msg(content.label) + '</span></a>';
	    		                             li = document.createElement("li");
	    		                             li.innerHTML = html;
	    		                             menuItem = new YAHOO.widget.MenuItem(li, config);

	    		                             menuItems.push(menuItem);
	    		             
	    		               		  }
    		                     }
    		                     //Added
    		                     
    		                     if (menuItems.length == 0)
    		                     {
    		                        menuItems.push(this.msg("label.empty"));
    		                     }
    		                     templateNodesMenu.clearContent();
    		                     templateNodesMenu.addItems(menuItems);
    		                     templateNodesMenu.render(); 
    		                  },
    		                  scope: this
    		               }
    		            });
    		         }
    		            
    	     };
    	  }(Alfresco.DocListToolbar.prototype.onCreateByTemplateNodeBeforeShow, this.options.createContentActions));
    	    	  
      },
      
      newOfficeDocument: function rplp_onNewOfficeDocument(file, actionDefinition) {
    	  if ("true" == actionDefinition.newnamedialog) 
    	  {
	    	//Ask for new name ...
	    	var nameheader = this.msg("message.name.header");
	    	var nametext = this.msg("message.name.text");
	  	    Alfresco.util.PopupManager.getUserInput(
			         {
			            title: nameheader,
			            text: nametext,
			            input: "text",
			            value: actionDefinition.filename,
			            callback:
			            {
			               fn: function rplp_promptNewnameCallback (newNodeName, obj)
			               {
			            	   actionDefinition.filename = newNodeName;
			            	   
			            	   this._launch(file, actionDefinition);
			               },
			               obj: {},
			               scope: this
			            }
			         });
    	 } 
    	 else
	     {
    	 	this._launch(file, actionDefinition);
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
    	  var feedbackMessage = this.myWait();
    	  
    	  var me = this;
           Alfresco.util.Ajax.jsonPost(
                   {
                      url: Alfresco.constants.PROXY_URI + "msoffice-creatable/copyin",
                      dataObj:
                      {
                          sourcePath: sourcePath,
                          parentNodeRef: newParentNodeRef
                      },
                      failureCallback: { fn: function (response)
                          { feedbackMessage.destroy(); } },
                      successCallback:
                      {
                    	  
                         fn: function (response)
                         {
                            if (null != response && null != response.json) {
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
	                                
	                                feedbackMessage.destroy();
	                         	
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
                      //successMessage: this.msg("message.create-content-by-template-node.success", sourcePath),
                      failureMessage: this.msg("message.create-content-by-template-node.failure", sourcePath)
                   });
		 
      },
      
      _launch: function dla_launch(file, actionDefinition) {
    	  if (this._launchOnlineCreator(file, actionDefinition))
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
    	   var feedbackMessage = this.myWait();

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

          var returnValue = false;
          if (YAHOO.env.ua.ie > 0)
          {
        	  returnValue = this._launchOnlineCreatorIE(controlProgID, record, appProgID);
          }
          
          if (Alfresco.util.isSharePointPluginInstalled())
          {
        	  returnValue = this._launchOnlineCreatorPlugin(record, appProgID);
          }
          else
          {
        	 feedbackMessage.destroy();
             Alfresco.util.PopupManager.displayPrompt(
             {
                text: this.msg("actions.editOnline.failure", loc.file)
             });
          }
          
          feedbackMessage.destroy();
          return returnValue;
          
       }

       // No success in launching application via ActiveX control; launch the WebDAV URL anyway
       return window.open(record.onlineEditUrl, "_blank");
    },

    myWait: function dla_myWait() {
    	
            var prompt = new YAHOO.widget.SimpleDialog("message",
                {
                   close: false,
                   constraintoviewport: true,
                   draggable: true,
                   effect: null,
                   modal: true,
                   visible: false,
                   zIndex: 100
                });

          // Show the prompt text
          prompt.setBody("<span class='wait'>" + this.msg("message.creating") + "</span>");

          // Add the dialog to the dom, center it and show it.
          prompt.render(document.body);
          prompt.center();
          Dom.get("message_h").style.display = 'none';
          prompt.show();
          return prompt;
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
  



      