<#escape x as jsonUtils.encodeJSONString(x)>
{
   "success": <#if name??>true<#else>false</#if>,
   "name": "${name?if_exists}"
}
</#escape>