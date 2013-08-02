<#escape x as jsonUtils.encodeJSONString(x)>
{
   "success": name??,
   "name": "${name?if_exists}"
}
</#escape>