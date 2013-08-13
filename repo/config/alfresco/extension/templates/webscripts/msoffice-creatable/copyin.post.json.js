/**
 * Document List Component: Create New Node - create copy of node template in the Data Dictionary
 */
function main()
{
   // get the arguments - expecting the "sourceNodeRef" and "parentNodeRef" of the source node to copy
   // and the parent node to contain the new copy of the source.
   var sourcePath = json.get("sourcePath");
   if (sourcePath == null || sourcePath.length === 0)
   {
      status.setCode(status.STATUS_BAD_REQUEST, "Mandatory 'sourcePath' parameter missing.");
      return;
   }
   var nodes = search.selectNodes(sourcePath);
   
   var parentNodeRef = json.get("parentNodeRef");
   if (parentNodeRef == null || parentNodeRef.length === 0)
   {
      status.setCode(status.STATUS_BAD_REQUEST, "Mandatory 'parentNodeRef' parameter missing.");
      return;
   }
   
   // get the nodes and perform the copy - permission failures etc. will produce a status code response
   var sourceNode = null;
   var parentNode = null;
   if (null != nodes && 0 < nodes.length) 
   {
	   var sourceNode = nodes[0];//search.findNode(sourceNodeRef),
	   if (null != sourceNode) {
		   parentNode = search.findNode(parentNodeRef);
		   
	   }
   }
   if (sourceNode == null || parentNode == null)
   {
      status.setCode(status.STATUS_NOT_FOUND, "Source or destination node is missing for copy operation.");
   } else {
	   var newname = json.get("newname");
	   var newnode = sourceNode.copy(parentNode);
	   
	   if (null != newnode) {
		   var cont = 10;
		   while (0 < cont) {
			   try {
				   newnode.properties.name = findGoodname(newname, parentNode.children);
				   newnode.save();
				   cont = 0;
			   } catch(e) {
				   newname = "copy-of-" + newname; 
				   cont--;
				   //hmmm
				   var newnode = sourceNode.copy(parentNode);
			   }
		   }
		   model.name = newnode.properties.name;
	   }
   }
}

function findGoodname(newname, children) {
	for (var child in children) {
		if (newname == ""+children[child].name) {
			return findGoodname("copy of " + newname, children);
		}
	}
	return newname;
}


main();