<#  ADquickSearch
    
    author: Sergey Trishkin
    mail:   grdvsng@gmail.com 

    description: 
        Module for quick data search in the Active Directory as a data-table.
    
    Example:
        [ADquickSearch]$ADSearch = [ADquickSearch]::new();
        [object]$request         = $ADSearch.request('displayName', 'sAMAccountName = "SpiderMan"');

        write-host $request;
        >>  "Peter Parker"

#>

class ADquickSearch
{
    [object] $ADconnection;
    [object] $Commander;
    [string] $rootDomainNamingContext;

    hidden [object] generateCollectioFromFields([__ComObject]$objRecordSet, [array]$keys)
    {
        [object]$request = @{};
        [int]$n          = -1;

        forEach ($field in $objRecordSet.Fields)
        {
            $request[$keys[$n]] = $field.Value;
            $n += 1;
        }

        return $request;
    }

    [object] request([string]$selector, [string]$sqlCondition)
    {
        <#  Class ADquickSearch method requests.]
                selector     - String - select condition;
                sqlCondition - String - condition for searching.
        
            rerturn:
                if   1 selector return String;
                else return hash table like @{selector1: value, selector2: value, ...}

        #>

        [string]$this.Commander.CommandText = "SELECT " + $selector + " FROM 'LDAP://" + $this.rootDomainNamingContext +"' WHERE " + $sqlCondition;
        [object]$objRecordSet               = $this.Commander.Execute();
        [array]$keys                        = $selector -split ", ";
        $result                             = $null;
        
        if ($objRecordSet.Fields.Count -eq 1) {
            return $objRecordSet.Fields(0).Value
        } else {
            return $this.generateCollectioFromFields($objRecordSet, $keys);
        }
    }

    [void] close()
    {
        [void]$this.ADconnection.Close();
    }

    [void] connect()
    {
        [void]$this.ADconnection.Open( "Active Directory Provider");
        
        [object]$this.Commander.ActiveConnection                  = $this.ADconnection;
        [int]$this.Commander.Properties.item("Searchscope").value = 2; #subTree
    }

    ADquickSearch()
    {
       $this.ADconnection            = new-Object -com "ADODB.Connection";
       $this.Commander               = new-Object -com "ADODB.Command";
       $this.rootDomainNamingContext = ([ADSI]"LDAP://RootDSE").rootDomainNamingContext;   
       $this.ADconnection.Provider   = "ADsDSOObject";

       $this.connect();   
    }
}