# ADquickSearch
  ### author: Sergey Trishkin
  ### mail:   grdvsng@gmail.com 

  ### description: 
  > #### Module for quick data search in the Active Directory as a data-table.
    
  ###Example:
```PowerShell
[ADquickSearch]$ADSearch = [ADquickSearch]::new();
[object]$request         = $ADSearch.request('displayName', 'sAMAccountName = "SpiderMan"');

write-host $request;
> "Peter Parker"
  
```
