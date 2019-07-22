<#  ADquickSearch
    
    author: Sergey Trishkin
    mail:   grdvsng@gmail.com 

    description: 
        Module for quick data search in the Active Directory as a data-table.
    
    Example:
        [ADquickSearch]$ADSearch = [ADquickSearch]::new();
        [object]$request         = $ADSearch.request('displayName', 'sAMAccountName = "SpiderMan"');

        write-host $request.rows.displayName;
        >>  "Peter Parker"

#>

function ADTableConnection([string]$selector, [string]$sqlCondition)
{
    [object] $ADconnection;
    [object] $Commander;
    [string] $rootDomainNamingContext;

    function getAttsFromField([object]$fields)
    {
        [object]$row = @{};
        
        forEach ($field in $fields)
        {
            $row[$field.Name] = $field.Value;
        }

        return $row;
    }

    function generateCollectioFromFields([__ComObject]$objRecordSet)
    {
        [hashtable]$diction  = @{};
        [array]$diction.rows = @();
        
        while (!$objRecordSet.EOF)
        {
            $diction.rows += getAttsFromField $objRecordSet.Fields;
            $objRecordSet.MoveNext();
        }
        
        return $diction;

    }

    function requests([object]$ADconnection, [object]$Commander, [string]$rootDomainNamingContext, [string]$selector, [string]$sqlCondition)
    {
        [string]$Commander.CommandText = "SELECT " + $selector + " FROM 'LDAP://" + $rootDomainNamingContext +"' WHERE " + $sqlCondition;
        [object]$objRecordSet          = $Commander.Execute();
        [hashtable]$result             = generateCollectioFromFields $objRecordSet;
        
        $objRecordSet.Close();
        return $result;
    } 

    function connect([object]$ADconnection, [object]$Commander)
    {
        $ADconnection.Open( "Active Directory Provider");
        
        [object]$Commander.ActiveConnection                  = $ADconnection;
        [int]$Commander.Properties.item("Searchscope").value = 2; #subTree
    }

    $ADconnection            = new-Object -com "ADODB.Connection";
    $Commander               = new-Object -com "ADODB.Command";
    $rootDomainNamingContext = ([ADSI]"LDAP://RootDSE").rootDomainNamingContext;   
    $ADconnection.Provider   = "ADsDSOObject";

    connect $ADconnection $Commander;
    
    [array]$result = requests $ADconnection $Commander $rootDomainNamingContext $selector $sqlCondition;
    return $result;
}

function getValue([string]$value, [bool]$tryingBefore)
{
    [string]$selector      = 'mail, telephoneNumber, sAMAccountName, displayName, postalAddress';
    [string]$condition     = "mail = '{0}' or telephoneNumber = '{0}' or sAMAccountName = '{0}' or displayName = '{0}' or sn = '{0}'" -f $value;
    [object]$curentdiction = ADTableConnection $selector $condition;
   
    if (!$tryingBefore -and (!$condition.rows)) {
        return getValue ("{0}*" -f $value) $true;
    } 

    return $curentdiction;

}
