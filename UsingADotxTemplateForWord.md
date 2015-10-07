# Using a .dotx Template for Word #

You can today use a .dotx template, you will have to provide document and override default configuration for it. I have put some documentation in the configuration

See http://code.google.com/p/alfresco-msoffice-creatable/source/browse/share/config/alfresco/web-extension/site-data/extensions/msoffice-creatable.xml

In your share-config-custom.xml: you can add something like:

```
 <!-- Document Library config section -->
   <config evaluator="string-compare" condition="DocumentLibrary" replace="true">   
   <!--  MS Office - create new document integration  -->
      <msoffice-creatable>
         <creatable-types>            
<!--  use a template ...  -->
            <creatable type="doc" index="2001">site-54/documentLibrary/template1.dotx</creatable>
         </creatable-types>
      </msoffice-creatable>
      </config>
```

This will try to use a dotx file from 'site-54' in 'documentLibrary' called 'template1.dotx', so this file must be there and readable by the users.